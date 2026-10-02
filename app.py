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
        "maizena", "detalhe", "obs", "registrado_em", "enfardadora", "silo_ini", "silo_fim"]
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
                "Maizena", "Nível silo (início)", "Nível silo (fim)", "Lote", "Observações"]


def _linha_leitura(row: dict):
    return [
        row.get("data", ""), row.get("hora_ini", ""), row.get("hora_fim", ""), row.get("etapa", ""),
        row.get("enfardadora", ""), row.get("tipo", ""), row.get("aspereza", ""),
        row.get("vazao_tph", ""), row.get("pct_nominal", ""), row.get("parado_min", ""),
        row.get("motivo", ""), row.get("maizena", ""), row.get("silo_ini", ""),
        row.get("silo_fim", ""), row.get("lote", ""), row.get("obs", ""),
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


def salvar(row: dict):
    row["registrado_em"] = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
    vals = [row.get(c, "") for c in COLS]
    if usa_sheets():
        get_ws().append_row(vals, value_input_option="USER_ENTERED")
        get_ws_leitura().append_row(_linha_leitura(row), value_input_option="USER_ENTERED")
    else:
        df = pd.DataFrame([vals], columns=COLS)
        df.to_csv(CSV_LOCAL, mode="a", header=not CSV_LOCAL.exists(), index=False)


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


def paradas(p, titulo="Minutos parado"):
    """Minutos + caixinhas de motivo. Retorna (minutos, texto do motivo)."""
    par = st.number_input(titulo, 0, 600, 0, key=p + "_par")
    if par == 0:
        return 0, ""
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
    return par, " ; ".join(mots)


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


# ---------- pré-bene e bene (ciclo: hora do 1º e 2º fechamento, por saída) ----------
def tela_bags(etapa, chave, nomes, maizena=False, nota=""):
    st.subheader(etapa)
    if nota:
        st.caption(nota)
    data = st.date_input("Data", datetime.now(TZ).date(), key=chave + "_data")
    lote = st.text_input("Lote (opcional)", key=chave + "_lote")
    c1, c2 = st.columns(2)
    tipo = c1.selectbox("Tipo de feijão", TIPOS, key=chave + "_tipo")
    asp = c2.radio("Aspereza", ["Liso", "Áspero"], horizontal=True, key=chave + "_asp")
    with st.expander("Peso do bag"):
        peso = st.number_input("Peso do bag (kg)", 100, 2000, 900, 10, key=chave + "_peso")

    extra_key = chave + "_extras"
    if extra_key not in st.session_state:
        st.session_state[extra_key] = []
    c1, c2 = st.columns([3, 1])
    c1.markdown("**Ciclo de cada saída**")
    if c2.button("+ Adicionar saída", key=chave + "_addbtn", use_container_width=True):
        st.session_state[extra_key].append(f"Saída extra {len(st.session_state[extra_key]) + 1}")

    for idx, nome_extra in enumerate(list(st.session_state[extra_key])):
        c1, c2 = st.columns([4, 1])
        novo_nome = c1.text_input("Nome da saída extra", nome_extra,
                                  key=f"{chave}_extranome{idx}")
        st.session_state[extra_key][idx] = novo_nome
        if c2.button("Remover", key=f"{chave}_extrarm{idx}"):
            st.session_state[extra_key].pop(idx)
            st.rerun()

    rows = []
    for nome in nomes + st.session_state[extra_key]:
        with st.expander(nome, expanded=(nome == nomes[0])):
            medir = st.checkbox(f"Medir {nome} agora", key=f"{chave}_{nome}_on",
                                value=(nome == nomes[0]))
            if not medir:
                continue
            ini = campo_hora("1º fechamento", f"{chave}_{nome}_ini")
            fim = campo_hora("2º fechamento", f"{chave}_{nome}_fim")
            dur = minutos(ini, fim)
            if dur <= 0:
                st.error("O 2º fechamento precisa ser depois do 1º.")
                continue
            st.caption(f"Ciclo: {dur // 60}h{dur % 60:02d} ({dur} min)")
            v = peso / 1000 / (dur / 60)
            st.metric(f"Vazão — {nome}", f"{v:.3f} ton/h")
            extra = {"detalhe": f"{nome}: 1 bag de {peso}kg em {dur} min"}
            if maizena and nome == nomes[0]:
                extra["maizena"] = st.radio("Maizena usada? (opcional, se souber)",
                                            ["Não sei", "Não", "Sim"], horizontal=True,
                                            key=f"{chave}_{nome}_mz")
            rows.append({"etapa": f"{etapa} — {nome}", "data": str(data), "lote": lote,
                        "tipo": tipo, "aspereza": asp, "hora_ini": ini.strftime("%H:%M"),
                        "hora_fim": fim.strftime("%H:%M"), "duracao_min": dur, "parado_min": 0,
                        "kg": peso, "vazao_tph": round(v, 3), "pct_nominal": "", **extra})
    par, mot = paradas(chave)
    obs = st.text_area("Observações", key=chave + "_obs")
    for r in rows:
        r["motivo"] = mot
        r["obs"] = obs
        if par:
            r["parado_min"] = par
    confirmar(rows)


# ---------- embaladora + 2 enfardadoras ----------
def tela_embalagem():
    st.subheader("Embaladora + enfardadoras")
    base, dur = cabecalho("emb", com_tipo=False)
    with st.expander("Peso por contagem do contador"):
        kgc = st.number_input("kg por unidade do contador", 0.1, 100.0, 1.0, 0.1, key="emb_kgc",
                              help="1 se o contador conta sacos de 1 kg; 30 se conta fardos.")
        peso_fardo = st.number_input("Peso do fardo (kg), para o contador da enfardadora",
                                     1, 200, 30, key="emb_pfardo")
    falta = st.number_input("Esperando feijão (min, geral)", 0, 600, 0, key="emb_falta")
    rows = []
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

        with st.expander(f"Contador da {enf.lower()} (fardos)", expanded=True):
            a, b = st.columns(2)
            e_ini = a.number_input("Contador início", 0, 100_000_000, 0, key=f"{k}_eci")
            e_fim = b.number_input("Contador fim", 0, 100_000_000, 0, key=f"{k}_ecf")
            kg_enf = max(e_fim - e_ini, 0) * peso_fardo
            if kg_enf > 0 and dur > 0:
                st.caption(f"{enf}: {kg_enf / 1000 / (dur / 60):.2f} ton/h (janela cheia, sem descontar paradas)")

        kg, det, par_tot, mots = 0.0, [], 0, []
        for i in maqs:
            with st.expander(f"Máquina {i}", expanded=True):
                a, b = st.columns(2)
                ini = a.number_input("Contador início", 0, 100_000_000, 0, key=f"emb_ci{i}")
                fim = b.number_input("Contador fim", 0, 100_000_000, 0, key=f"emb_cf{i}")
                pm, mm = paradas(f"emb_m{i}", "Min parada da máquina")
                q = max(fim - ini, 0) * kgc
                kg += q
                par_tot += pm
                if q > 0 and dur - pm > 0:
                    st.caption(f"Máquina {i}: {q / 1000 / ((dur - pm) / 60):.2f} ton/h (rodando {dur - pm} min)")
                det.append(f"M{i}:{int(q)}kg,par{pm}min")
                if mm:
                    mots.append(f"M{i}: {mm}")
        with st.expander(f"{enf} (parada)", expanded=True):
            pe, me = paradas(f"emb_{k}_enf", "Min parada da enfardadora")
            fardos_enf = max(e_fim - e_ini, 0)
            pal = fardos_enf / 35  # 1 palete = 1.050 kg = 35 fardos de 30 kg
            if fardos_enf > 0:
                st.caption(f"{fardos_enf} fardos ≈ {pal:.2f} paletes ({kg_enf / 1000:.2f} t) "
                          f"| Soma máquinas: {kg / 1000:.2f} t")
        if me:
            mots.append(f"Enfardadora: {me}")
        par_tot += pe
        # usa o contador da própria enfardadora quando preenchido; senão, soma das máquinas
        kg_final = kg_enf if kg_enf > 0 else kg
        r = resultado({**base, "etapa": "Embaladora", "tipo": tipo, "aspereza": asp}, kg_final, dur, 0, 3.0,
                      {"enfardadora": enf, "silo_ini": silo_ini, "silo_fim": silo_fim,
                       "motivo": " ; ".join(mots),
                       "detalhe": "; ".join(det) + f"; fardos_enf:{fardos_enf}; paletes_calc:{pal:.2f}; "
                                                    f"contador_enf_kg:{kg_enf:.0f}"})
        if r:
            r["parado_min"] = par_tot + falta
            rows.append(r)
    if falta:
        for r in rows:
            r["motivo"] = (r["motivo"] + " ; " if r["motivo"] else "") + "Falta de feijão"
    obs = st.text_area("Observações", key="emb_obs")
    for r in rows:
        r["obs"] = obs
    confirmar(rows)


# ---------- gráficos ----------
def tela_graficos():
    st.subheader("Comparação áspero × liso")
    df = carregar()
    if df.empty:
        st.info("Ainda não há medições salvas.")
        return
    df = df.dropna(subset=["vazao_tph"])
    sel = st.selectbox("Tipo de feijão", ["Todos"] + sorted(df["tipo"].dropna().unique().tolist()))
    if sel != "Todos":
        df = df[df["tipo"] == sel]
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
    p = df[(df["parado_min"] > 0) & (df["motivo"].astype(str) != "")].copy()
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
