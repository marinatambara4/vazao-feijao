import re
from datetime import datetime, time
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Vazão feijão", page_icon="🫘", layout="centered")
TZ = ZoneInfo("America/Sao_Paulo")
CSV_LOCAL = Path("medicoes.csv")
# colunas novas ficam no fim para não bagunçar linhas antigas
COLS = ["etapa", "data", "hora_ini", "hora_fim", "tipo", "aspereza", "lote",
        "duracao_min", "parado_min", "motivo", "kg", "vazao_tph", "pct_nominal",
        "maizena", "detalhe", "obs", "registrado_em", "enfardadora", "silo_ini", "silo_fim",
        "id_medicao", "status_medicao", "modo_medicao"]
NIVEIS_SILO = ["Funil", "Anel 1", "Anel 2", "Anel 3"]
SILO_OPCOES = NIVEIS_SILO + ["Não foi possível medir"]
TIPOS = ["Preto", "Carioca", "Vermelho"]


# ---------- armazenamento (Google Sheets ou CSV local) ----------
def usa_sheets():
    return "gcp_service_account" in st.secrets and "sheet_id" in st.secrets


@st.cache_resource
def get_ws():
    import gspread
    from google.oauth2.service_account import Credentials
    info = dict(st.secrets["gcp_service_account"])
    info["private_key"] = info["private_key"].replace("\\n", "\n").strip()
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    sh = gspread.authorize(creds).open_by_key(st.secrets["sheet_id"])
    try:
        ws = sh.worksheet("medicoes")
    except Exception:
        ws = sh.add_worksheet("medicoes", rows=1000, cols=len(COLS))
    if ws.col_count < len(COLS):
        ws.resize(cols=len(COLS))
    if ws.row_values(1) != COLS:
        ws.update(range_name="A1", values=[COLS])
    return ws


LEITURA_COLS = ["Data", "Hora início", "Hora fim", "Etapa", "Enfardadora", "Tipo de feijão",
                "Aspereza", "Vazão (ton/h)", "% da referência", "Min. parado", "Motivo da parada",
                "Maizena", "Nível silo (início)", "Nível silo (fim)", "Lote", "Observações", "Status"]


def _linha_leitura(row: dict):
    return [
        row.get("data", ""), row.get("hora_ini", ""), row.get("hora_fim", ""), row.get("etapa", ""),
        row.get("enfardadora", ""), row.get("tipo", ""), row.get("aspereza", ""),
        row.get("vazao_tph", ""), row.get("pct_nominal", ""), row.get("parado_min", ""),
        row.get("motivo", ""), row.get("maizena", ""), row.get("silo_ini", ""),
        row.get("silo_fim", ""), row.get("lote", ""), row.get("obs", ""),
        row.get("status_medicao", "Concluída"),
    ]


@st.cache_resource
def get_ws_leitura():
    ws = get_ws()  # garante que a planilha já foi aberta
    sh = ws.spreadsheet
    try:
        wl = sh.worksheet("leitura")
    except Exception:
        wl = sh.add_worksheet("leitura", rows=1000, cols=len(LEITURA_COLS))
    if wl.row_values(1) != LEITURA_COLS:
        wl.update(range_name="A1", values=[LEITURA_COLS])
    return wl


def _append_medicoes(row: dict):
    vals = [row.get(c, "") for c in COLS]
    if usa_sheets():
        get_ws().append_row(vals, value_input_option="USER_ENTERED")
    else:
        df = pd.DataFrame([vals], columns=COLS)
        df.to_csv(CSV_LOCAL, mode="a", header=not CSV_LOCAL.exists(), index=False)


def _append_leitura(row: dict):
    if usa_sheets():
        get_ws_leitura().append_row(_linha_leitura(row), value_input_option="USER_ENTERED")
    # sem Sheets, os dados já estão completos no CSV; não há aba de leitura separada


def salvar(row: dict):
    """Salva uma medição já concluída de uma vez (pré-bene e bene continuam usando isso)."""
    row["registrado_em"] = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
    row.setdefault("status_medicao", "Concluída")
    row.setdefault("modo_medicao", "")
    row.setdefault("id_medicao", "")
    _append_medicoes(row)
    _append_leitura(row)


def novo_id(prefixo):
    return f"{prefixo}_{datetime.now(TZ).strftime('%Y%m%d%H%M%S%f')}"


def pendentes(key_prefixo):
    """Medições iniciadas e ainda não encerradas para esse equipamento."""
    df = carregar()
    if df.empty or "id_medicao" not in df.columns:
        return pd.DataFrame()
    col_id = df["id_medicao"].astype(str)
    mask = col_id.str.startswith(key_prefixo + "_") & (df.get("status_medicao", "") == "Em andamento")
    return df[mask]


def atualizar_medicao(id_medicao: str, updates: dict):
    """Encontra a linha pelo id e atualiza só os campos passados. Retorna a linha final ou None."""
    if usa_sheets():
        ws = get_ws()
        try:
            cell = ws.find(id_medicao)
        except Exception:
            return None
        if not cell:
            return None
        header = ws.row_values(1)
        atuais = ws.row_values(cell.row)
        full = {h: (atuais[i] if i < len(atuais) else "") for i, h in enumerate(header)}
        full.update(updates)
        vals = [full.get(c, "") for c in COLS]
        ws.update(f"A{cell.row}", [vals], value_input_option="USER_ENTERED")
        return full
    else:
        if not CSV_LOCAL.exists():
            return None
        df = pd.read_csv(CSV_LOCAL, dtype=str).fillna("")
        if "id_medicao" not in df.columns:
            return None
        idx_list = df.index[df["id_medicao"].astype(str) == id_medicao].tolist()
        if not idx_list:
            return None
        idx = idx_list[0]
        for k, v in updates.items():
            df.loc[idx, k] = v
        df.to_csv(CSV_LOCAL, index=False)
        return df.loc[idx].to_dict()


def carregar() -> pd.DataFrame:
    if usa_sheets():
        df = pd.DataFrame(get_ws().get_all_records())
    elif CSV_LOCAL.exists():
        df = pd.read_csv(CSV_LOCAL)
    else:
        return pd.DataFrame(columns=COLS)
    for c in ["vazao_tph", "pct_nominal", "parado_min", "duracao_min", "kg"]:
        if c in df:
            df[c] = pd.to_numeric(df[c].astype(str).str.replace(",", "."), errors="coerce")
    return df


# ---------- componentes ----------
def agora():
    return datetime.now(TZ).time().replace(second=0, microsecond=0)


def campo_hora(label, key):
    c1, c2 = st.columns([3, 2])
    if c2.button("Agora", key=key + "_b", use_container_width=True):
        st.session_state[key] = agora()
    return c1.time_input(label, value=st.session_state.get(key, time(9, 0)), key=key, step=60)


def minutos(ini: time, fim: time) -> int:
    return (fim.hour * 60 + fim.minute) - (ini.hour * 60 + ini.minute)


def cabecalho(p, com_tipo=True):
    data = st.date_input("Data", datetime.now(TZ).date(), key=p + "_data")
    ini = campo_hora("Hora início", p + "_ini")
    fim = campo_hora("Hora fim", p + "_fim")
    dur = minutos(ini, fim)
    if dur <= 0:
        st.error("A hora fim precisa ser depois da hora início.")
    else:
        st.caption(f"Janela: {dur // 60}h{dur % 60:02d} ({dur} min)")
    base = dict(data=str(data), hora_ini=ini.strftime("%H:%M"), hora_fim=fim.strftime("%H:%M"),
                duracao_min=dur, lote=st.text_input("Lote (opcional)", key=p + "_lote"))
    if com_tipo:
        c1, c2 = st.columns(2)
        base["tipo"] = c1.selectbox("Tipo de feijão", TIPOS, key=p + "_tipo")
        base["aspereza"] = c2.radio("Aspereza", ["Liso", "Áspero"], horizontal=True, key=p + "_asp")
    return base, dur


def caixas_motivo(p):
    """Caixinhas de motivo de parada. Retorna o texto combinado."""
    st.caption("Motivo da parada")
    c1, c2 = st.columns(2)
    asp = c1.checkbox("Aspereza do grão", key=p + "_m1")
    ele = c2.checkbox("Pane elétrica", key=p + "_m2")
    meca = c1.checkbox("Pane mecânica", key=p + "_m3")
    hid = c2.checkbox("Pane hidráulica", key=p + "_m4")
    maq = st.text_input("Qual máquina?", key=p + "_maq") if (ele or meca or hid) else ""
    outro = st.checkbox("Outro", key=p + "_m5")
    txt = st.text_input("Especifique", key=p + "_outro") if outro else ""
    suf = f" ({maq})" if maq else ""
    mots = ["Aspereza do grão"] if asp else []
    mots += [n + suf for ok, n in [(ele, "Pane elétrica"), (meca, "Pane mecânica"),
                                   (hid, "Pane hidráulica")] if ok]
    if outro:
        mots.append(f"Outro: {txt}" if txt else "Outro")
    return " ; ".join(mots)


def paradas(p, titulo="Minutos parado"):
    """Minutos + caixinhas de motivo. Retorna (minutos, texto do motivo)."""
    par = st.number_input(titulo, 0, 600, 0, key=p + "_par")
    if par == 0:
        return 0, ""
    return par, caixas_motivo(p)


def bloco_equipamento(nome, key, kg_por_unidade, nominal, etapa_nome, data, tipo, aspereza, obs,
                      enfardadora="", silo_ini="", silo_fim="", mostrar_paletes=False):
    """Equipamento com hora própria. Fluxo em 2 passos, cada um salvo na hora:
    1) 'Iniciar' grava hora/contador início (ou início de parada) direto na planilha.
    2) Depois, 'Encerrar' completa com hora/contador fim e calcula a vazão.
    Assim nada se perde se você trocar de aba ou a conexão cair no meio."""
    st.markdown(f"**{nome}**")
    pend = pendentes(key)

    if not pend.empty:
        linha = pend.sort_values("registrado_em", ascending=False).iloc[0]
        modo = linha.get("modo_medicao", "Rodando") or "Rodando"
        st.warning(f"Medição em andamento desde {linha['hora_ini']} ({modo}). "
                   f"Encerre antes de iniciar outra.")
        fim = campo_hora("Hora fim", key + "_fimresume")
        try:
            ini_time = datetime.strptime(str(linha["hora_ini"]), "%H:%M").time()
        except Exception:
            ini_time = time(0, 0)
        dur = minutos(ini_time, fim)
        if dur <= 0:
            st.error("A hora fim precisa ser depois do início registrado.")
            return
        st.caption(f"Desde o início: {dur // 60}h{dur % 60:02d} ({dur} min)")

        if modo == "Rodando":
            cf = st.number_input("Contador fim", 0, 100_000_000, 0, key=key + "_cfresume")
            if st.button(f"Encerrar {nome}", key=key + "_encerrar", type="primary"):
                det = str(linha.get("detalhe", ""))
                ci = 0.0
                if "contagem_ini:" in det:
                    try:
                        ci = float(det.split("contagem_ini:")[1].split(";")[0])
                    except Exception:
                        ci = 0.0
                qtd = max(cf - ci, 0)
                kg = qtd * kg_por_unidade
                v = kg / 1000 / (dur / 60) if kg > 0 else 0
                pct = v / nominal * 100 if kg > 0 else 0
                updates = {"hora_fim": fim.strftime("%H:%M"), "duracao_min": dur, "parado_min": 0,
                          "kg": round(kg, 1), "vazao_tph": round(v, 3) if kg > 0 else "",
                          "pct_nominal": round(pct, 1) if kg > 0 else "",
                          "status_medicao": "Concluída", "detalhe": f"contagem:{qtd}", "obs": obs}
                full = atualizar_medicao(linha["id_medicao"], updates)
                if full:
                    _append_leitura(full)
                    msg = f"Medição encerrada: {v:.2f} ton/h." if kg > 0 else "Medição encerrada."
                    if mostrar_paletes and kg > 0:
                        msg += f" ≈ {kg / 1050:.2f} paletes."
                    st.success(msg)
                    st.rerun()
                else:
                    st.error("Não consegui encontrar essa medição para encerrar. Tente de novo.")
        else:
            if st.button(f"Encerrar parada — {nome}", key=key + "_encerrarpar", type="primary"):
                updates = {"hora_fim": fim.strftime("%H:%M"), "duracao_min": dur, "parado_min": dur,
                          "status_medicao": "Concluída", "obs": obs}
                full = atualizar_medicao(linha["id_medicao"], updates)
                if full:
                    _append_leitura(full)
                    st.success("Parada encerrada e salva.")
                    st.rerun()
                else:
                    st.error("Não consegui encontrar essa medição para encerrar. Tente de novo.")
        return

    status = st.radio("Status agora", ["Rodando", "Parada"], horizontal=True, key=key + "_status")
    ini = campo_hora("Hora início", key + "_ini")
    if status == "Rodando":
        ci = st.number_input("Contador início (o que já está no visor agora)", 0, 100_000_000, 0,
                             key=key + "_ci")
        if st.button(f"Iniciar {nome} (salva já)", key=key + "_iniciar", type="primary"):
            row = {"id_medicao": novo_id(key), "status_medicao": "Em andamento", "modo_medicao": "Rodando",
                  "etapa": etapa_nome, "data": str(data), "hora_ini": ini.strftime("%H:%M"), "hora_fim": "",
                  "duracao_min": "", "parado_min": "", "kg": "", "vazao_tph": "", "pct_nominal": "",
                  "motivo": "", "detalhe": f"contagem_ini:{ci}", "tipo": tipo, "aspereza": aspereza,
                  "lote": "", "maizena": "", "enfardadora": enfardadora, "silo_ini": silo_ini,
                  "silo_fim": silo_fim, "obs": "",
                  "registrado_em": datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")}
            _append_medicoes(row)
            st.success(f"Início salvo às {ini.strftime('%H:%M')}. Volte aqui mais tarde para encerrar.")
            st.rerun()
    else:
        motivo = caixas_motivo(key)
        if st.button(f"Registrar parada de {nome} (salva já)", key=key + "_iniciarpar", type="primary"):
            row = {"id_medicao": novo_id(key), "status_medicao": "Em andamento", "modo_medicao": "Parada",
                  "etapa": etapa_nome, "data": str(data), "hora_ini": ini.strftime("%H:%M"), "hora_fim": "",
                  "duracao_min": "", "parado_min": "", "kg": "", "vazao_tph": "", "pct_nominal": "",
                  "motivo": motivo, "detalhe": "parada", "tipo": tipo, "aspereza": aspereza,
                  "lote": "", "maizena": "", "enfardadora": enfardadora, "silo_ini": silo_ini,
                  "silo_fim": silo_fim, "obs": "",
                  "registrado_em": datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")}
            _append_medicoes(row)
            st.success(f"Início da parada salvo às {ini.strftime('%H:%M')}.")
            st.rerun()



def resultado(base, kg, dur, parado, nominal, extra):
    liq = dur - parado
    if liq <= 0 or kg <= 0:
        st.warning("Preencha bags/contadores e horários para ver a vazão.")
        return None
    v = kg / 1000 / (liq / 60)
    pct = v / nominal * 100
    c1, c2 = st.columns(2)
    c1.metric("Vazão real", f"{v:.2f} ton/h")
    c2.metric("% da referência", f"{pct:.0f}%", help=f"Referência: {nominal} ton/h")
    return {**base, "parado_min": parado, "kg": round(kg, 1), "vazao_tph": round(v, 3),
            "pct_nominal": round(pct, 1), **extra}


def confirmar(rows):
    rows = [r for r in rows if r]
    if not rows:
        return
    if st.button(f"Salvar medição ({len(rows)} registro{'s' if len(rows) > 1 else ''})",
                 type="primary", use_container_width=True):
        try:
            for r in rows:
                salvar(r)
            st.success("Medição salva.")
        except Exception as e:
            st.error(f"Não salvou: {e}")


# ---------- pré-bene e bene (janela de tempo + quantos bags fecharam) ----------
def tela_bags(etapa, chave, nomes, nominal=None, maizena=False, nota=""):
    st.subheader(etapa)
    if nota:
        st.caption(nota)
    base, dur = cabecalho(chave)
    with st.expander("Peso do bag"):
        peso = st.number_input("Peso do bag (kg)", 100, 2000, 900, 10, key=chave + "_peso")

    extra_key = chave + "_extras"
    if extra_key not in st.session_state:
        st.session_state[extra_key] = []
    c1, c2 = st.columns([3, 1])
    c1.markdown("**Bags por saída, no intervalo acima**")
    if c2.button("+ Adicionar saída", key=chave + "_addbtn", use_container_width=True):
        st.session_state[extra_key].append(f"Saída extra {len(st.session_state[extra_key]) + 1}")
    st.caption("Pode usar número quebrado para bag incompleto: 1/2 = 0,5 · 2/3 ≈ 0,67 · 3/4 = 0,75")

    for idx, nome_extra in enumerate(list(st.session_state[extra_key])):
        c1, c2 = st.columns([4, 1])
        novo_nome = c1.text_input("Nome da saída extra", nome_extra, key=f"{chave}_extranome{idx}")
        st.session_state[extra_key][idx] = novo_nome
        if c2.button("Remover", key=f"{chave}_extrarm{idx}"):
            st.session_state[extra_key].pop(idx)
            st.rerun()

    todas_saidas = nomes + st.session_state[extra_key]
    kgs = {}
    for nome in todas_saidas:
        kgs[nome] = st.number_input(nome, 0.0, 1000.0, 0.0, step=0.1, format="%.2f",
                                    key=f"{chave}_{nome}_bags") * peso

    par, mot = paradas(chave)
    maiz = None
    if maizena:
        maiz = st.radio("Maizena usada? (opcional, se souber)", ["Não sei", "Não", "Sim"],
                        horizontal=True, key=chave + "_mz")
    obs = st.text_area("Observações", key=chave + "_obs")

    rows = []
    for nome in todas_saidas:
        if kgs[nome] <= 0:
            continue
        extra = {"detalhe": f"{nome}: {kgs[nome] / peso:.2f} bag(s) de {peso}kg"}
        if maiz is not None:
            extra["maizena"] = maiz
        r = resultado({**base, "etapa": f"{etapa} — {nome}"}, kgs[nome], dur, par,
                      nominal if nominal else 1.0, extra)
        if r:
            r["motivo"] = mot
            r["obs"] = obs
            rows.append(r)
    if not any(kgs[n] > 0 for n in todas_saidas):
        st.info("Preencha os bags de ao menos uma saída para ver a vazão.")
    confirmar(rows)


# ---------- embaladora + 2 enfardadoras ----------
def tela_embalagem():
    st.subheader("Embaladora + enfardadoras")
    st.caption("Cada máquina e cada enfardadora tem seu próprio horário, e cada uma se salva sozinha: "
               "assim que você tiver a leitura inicial, clique em **Iniciar** — isso já fica gravado "
               "na hora. Quando voltar lá com a leitura final, abra esta tela de novo (ela lembra o "
               "que ficou pendente) e clique em **Encerrar**.")
    data = st.date_input("Data", datetime.now(TZ).date(), key="emb_data")
    with st.expander("Pesos de referência"):
        kgc = st.number_input("kg por unidade do contador da máquina", 0.1, 100.0, 1.0, 0.1, key="emb_kgc",
                              help="1 se o contador conta sacos de 1 kg; 30 se conta fardos.")
        peso_fardo = st.number_input("Peso do fardo (kg)", 1, 200, 30, key="emb_pfardo")
    obs = st.text_area("Observações (vale para o que você encerrar agora nesta tela)", key="emb_obs")

    for enf, maqs, k in [("Enfardadora 1", (1, 2), "e1"), ("Enfardadora 2", (3, 4), "e2")]:
        st.divider()
        st.markdown(f"### {enf} · máquinas {maqs[0]} e {maqs[1]}")
        c1, c2 = st.columns(2)
        tipo = c1.selectbox("Tipo de feijão", TIPOS, key=f"{k}_tipo")
        asp = c2.radio("Aspereza", ["Liso", "Áspero"], horizontal=True, key=f"{k}_asp")
        st.markdown(f"**Nível do silo pulmão ({enf})**")
        c3, c4 = st.columns(2)
        silo_ini = c3.radio("No início", SILO_OPCOES, key=f"{k}_silo_i")
        silo_fim = c4.radio("No fim", SILO_OPCOES, key=f"{k}_silo_f")

        with st.expander(f"{enf} — medição", expanded=True):
            bloco_equipamento(enf, f"{k}_enf", peso_fardo, 3.0, "Embaladora", data, tipo, asp, obs,
                              enfardadora=enf, silo_ini=silo_ini, silo_fim=silo_fim, mostrar_paletes=True)

        for i in maqs:
            with st.expander(f"Máquina {i} — medição"):
                bloco_equipamento(f"Máquina {i}", f"m{i}", kgc, 1.5, f"Embaladora — Máquina {i}",
                                  data, tipo, asp, obs, enfardadora=enf)


# ---------- gráficos ----------
def tela_graficos():
    st.subheader("Comparação áspero × liso")
    df_tudo = carregar()
    if df_tudo.empty:
        st.info("Ainda não há medições salvas.")
        return
    sel = st.selectbox("Tipo de feijão", ["Todos"] + sorted(df_tudo["tipo"].dropna().unique().tolist()))
    if sel != "Todos":
        df_tudo = df_tudo[df_tudo["tipo"] == sel]
    df = df_tudo.dropna(subset=["vazao_tph"])  # só medições com vazão calculada
    cores = {"Liso": "#2a9d8f", "Áspero": "#c8553d"}

    st.markdown("**1. Vazão real por etapa**")
    st.plotly_chart(px.box(df, x="etapa", y="vazao_tph", color="aspereza", points="all",
                           color_discrete_map=cores, labels={"vazao_tph": "ton/h", "etapa": ""}),
                    use_container_width=True)

    st.markdown("**2. Perda de vazão do áspero vs. liso**")
    m = df.groupby(["etapa", "aspereza"])["vazao_tph"].mean().unstack()
    if {"Liso", "Áspero"} <= set(m.columns):
        m["perda_%"] = (1 - m["Áspero"] / m["Liso"]) * 100
        st.plotly_chart(px.bar(m.reset_index(), x="etapa", y="perda_%", text_auto=".1f",
                               labels={"perda_%": "% de perda", "etapa": ""}), use_container_width=True)
    else:
        st.caption("Precisa de medições nos dois grupos.")

    st.markdown("**3. Vazão ao longo do dia**")
    d = df.copy()
    d["hora"] = pd.to_numeric(d["hora_ini"].astype(str).str[:2], errors="coerce")
    st.plotly_chart(px.scatter(d, x="hora", y="vazao_tph", color="aspereza", symbol="etapa",
                               color_discrete_map=cores), use_container_width=True)

    st.markdown("**4. Com maizena × sem maizena**")
    mz = df[df.get("maizena", pd.Series(dtype=str)).isin(["Sim", "Não"])]
    if mz.empty:
        st.caption("Sem medições com a maizena registrada (Sim/Não).")
    else:
        st.plotly_chart(px.box(mz, x="maizena", y="vazao_tph", color="aspereza", points="all",
                               facet_col="etapa", color_discrete_map=cores), use_container_width=True)

    st.markdown("**5. Minutos parados por motivo**")
    p = df_tudo[(pd.to_numeric(df_tudo["parado_min"], errors="coerce").fillna(0) > 0)
                & (df_tudo["motivo"].astype(str) != "")].copy()
    p["parado_min"] = pd.to_numeric(p["parado_min"], errors="coerce").fillna(0)
    if not p.empty:
        p["lista"] = p["motivo"].astype(str).str.split(" ; ")
        p["parado_min"] = p["parado_min"] / p["lista"].str.len()
        p = p.explode("lista")
        p["lista"] = p["lista"].str.replace(r"^(M\d|Enfardadora):\s*", "", regex=True)
        p["lista"] = p["lista"].apply(lambda s: re.sub(r"\s*\(.*\)", "", s))
        pp = p.groupby("lista")["parado_min"].sum().reset_index().sort_values("parado_min")
        st.plotly_chart(px.bar(pp, x="parado_min", y="lista", orientation="h",
                               labels={"lista": "", "parado_min": "min (distribuído entre motivos)"}),
                        use_container_width=True)
    else:
        st.caption("Sem paradas registradas.")

    st.markdown("**6. Embaladora: vazão por nível do silo (início)**")
    e = df[df["etapa"] == "Embaladora"]
    e = e[e.get("silo_ini", "").isin(NIVEIS_SILO)] if "silo_ini" in e else e.iloc[0:0]
    if not e.empty:
        st.plotly_chart(px.box(e, x="silo_ini", y="vazao_tph", color="aspereza",
                               points="all", color_discrete_map=cores,
                               category_orders={"silo_ini": NIVEIS_SILO}), use_container_width=True)
    else:
        st.caption("Sem medições da embaladora com nível do silo.")

    st.markdown("**7. Quantas medições por grupo**")
    st.dataframe(df.groupby(["etapa", "aspereza"]).size().unstack(fill_value=0))
    st.caption("Meta: 5 ou mais medições em cada célula.")
    with st.expander("Dados brutos"):
        st.dataframe(df)


# ---------- navegação ----------
st.title("Vazão do feijão")
if not usa_sheets():
    st.warning("Google Sheets não configurado: salvando em CSV local (medicoes.csv).")
tela = st.radio("Etapa", ["Pré-bene", "Bene", "Embaladora", "Gráficos"], horizontal=True)
if tela == "Pré-bene":
    tela_bags("Pré-beneficiamento", "pre", ["Feijão tipo 1"], maizena=True,
              nota="É aqui que a maizena costuma ser adicionada.")
elif tela == "Bene":
    tela_bags("Beneficiamento", "bene", ["Prata", "Bandinha"], maizena=True,
              nota="A vazão do feijão principal (que vai direto ao silo) é vista depois, na tela da Embaladora. "
                   "Deixe marcado se também passa maizena aqui, caso aplicável.")
elif tela == "Embaladora":
    tela_embalagem()
else:
    tela_graficos()
