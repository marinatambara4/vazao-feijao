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

COLS = [
    "etapa", "data", "hora_ini", "hora_fim", "tipo", "aspereza", "lote",
    "duracao_min", "parado_min", "motivo", "kg", "vazao_tph", "pct_nominal",
    "maizena", "detalhe", "obs", "registrado_em", "enfardadora", "silo_ini", "silo_fim",
    "id_medicao", "status_medicao", "modo_medicao"
]
NIVEIS_SILO = ["Funil", "Anel 1", "Anel 2", "Anel 3"]
SILO_OPCOES = NIVEIS_SILO + ["Não foi possível medir"]
TIPOS = ["Preto", "Carioca", "Vermelho"]
ASPEREZA_OPCOES = ["Levemente Liso", "Levemente Áspero", "Liso", "Áspero"]


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
        info, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
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


LEITURA_COLS = [
    "Data", "Hora início", "Hora fim", "Etapa", "Enfardadora", "Tipo de feijão",
    "Aspereza", "Vazão (ton/h)", "% da referência", "Min. parado", "Motivo da parada",
    "Maizena", "Nível silo (início)", "Nível silo (fim)", "Lote", "Observações", "Status"
]


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
    ws = get_ws()
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


def salvar(row: dict):
    row["registrado_em"] = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
    row.setdefault("status_medicao", "Concluída")
    row.setdefault("modo_medicao", "")
    row.setdefault("id_medicao", "")
    _append_medicoes(row)
    _append_leitura(row)


def novo_id(prefixo):
    return f"{prefixo}_{datetime.now(TZ).strftime('%Y%m%d%H%M%S%f')}"


def pendentes(key_prefixo):
    df = carregar()
    if df.empty or "id_medicao" not in df.columns:
        return pd.DataFrame()
    col_id = df["id_medicao"].astype(str)
    mask = col_id.str.startswith(key_prefixo + "_") & (df.get("status_medicao", "") == "Em andamento")
    return df[mask]


def atualizar_medicao(id_medicao: str, updates: dict):
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
        valores = get_ws().get_all_values()
        if not valores or len(valores) < 2:
            df = pd.DataFrame(columns=COLS)
        else:
            header = list(valores[0])
            vistos, header_unico = {}, []
            for h in header:
                vistos[h] = vistos.get(h, 0) + 1
                header_unico.append(h if vistos[h] == 1 else f"{h}_{vistos[h]}")
            n = len(header_unico)
            linhas = [list(r[:n]) + [""] * (n - len(r)) for r in valores[1:]]
            df = pd.DataFrame(linhas, columns=header_unico)
            for c in COLS:
                if c not in df.columns:
                    df[c] = ""
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


def campo_hora(label, key, default_val=None):
    c1, c2 = st.columns([3, 2])
    if c2.button("Agora", key=key + "_b", use_container_width=True):
        st.session_state[key] = agora()
    val_padrao = default_val if default_val else st.session_state.get(key, time(9, 0))
    return c1.time_input(label, value=val_padrao, key=key, step=60)


def minutos(ini: time, fim: time) -> int:
    return (fim.hour * 60 + fim.minute) - (ini.hour * 60 + ini.minute)


def caixas_motivo(p):
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
    par = st.number_input(titulo, 0, 600, 0, key=p + "_par")
    if par == 0:
        return 0, ""
    return par, caixas_motivo(p)


def bloco_equipamento(nome, key, kg_por_unidade, nominal, etapa_nome, data, tipo, aspereza, obs,
                      enfardadora="", silo_ini="", silo_fim="", mostrar_paletes=False):
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
                updates = {
                    "hora_fim": fim.strftime("%H:%M"),
                    "duracao_min": dur,
                    "parado_min": 0,
                    "kg": round(kg, 1),
                    "vazao_tph": round(v, 3) if kg > 0 else "",
                    "pct_nominal": round(pct, 1) if kg > 0 else "",
                    "status_medicao": "Concluída",
                    "detalhe": f"contagem:{qtd}",
                    "obs": obs
                }
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
                updates = {
                    "hora_fim": fim.strftime("%H:%M"),
                    "duracao_min": dur,
                    "parado_min": dur,
                    "status_medicao": "Concluída",
                    "obs": obs
                }
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
        ci = st.number_input("Contador início (o que já está no visor agora)", 0, 100_000_000, 0, key=key + "_ci")
        if st.button(f"Iniciar {nome} (salva já)", key=key + "_iniciar", type="primary"):
            row = {
                "id_medicao": novo_id(key),
                "status_medicao": "Em andamento",
                "modo_medicao": "Rodando",
                "etapa": etapa_nome,
                "data": str(data),
                "hora_ini": ini.strftime("%H:%M"),
                "hora_fim": "",
                "duracao_min": "",
                "parado_min": "",
                "kg": "",
                "vazao_tph": "",
                "pct_nominal": "",
                "motivo": "",
                "detalhe": f"contagem_ini:{ci}",
                "tipo": tipo,
                "aspereza": aspereza,
                "lote": "",
                "maizena": "",
                "enfardadora": enfardadora,
                "silo_ini": silo_ini,
                "silo_fim": silo_fim,
                "obs": "",
                "registrado_em": datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
            }
            try:
                _append_medicoes(row)
                st.success(f"Início salvo às {ini.strftime('%H:%M')}. Volte aqui mais tarde para encerrar.")
                st.rerun()
            except Exception as e:
                st.error(f"Não consegui salvar o início: {e}")
    else:
        motivo = caixas_motivo(key)
        if st.button(f"Registrar parada de {nome} (salva já)", key=key + "_iniciarpar", type="primary"):
            row = {
                "id_medicao": novo_id(key),
                "status_medicao": "Em andamento",
                "modo_medicao": "Parada",
                "etapa": etapa_nome,
                "data": str(data),
                "hora_ini": ini.strftime("%H:%M"),
                "hora_fim": "",
                "duracao_min": "",
                "parado_min": "",
                "kg": "",
                "vazao_tph": "",
                "pct_nominal": "",
                "motivo": motivo,
                "detalhe": "parada",
                "tipo": tipo,
                "aspereza": aspereza,
                "lote": "",
                "maizena": "",
                "enfardadora": enfardadora,
                "silo_ini": silo_ini,
                "silo_fim": silo_fim,
                "obs": "",
                "registrado_em": datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
            }
            try:
                _append_medicoes(row)
                st.success(f"Início da parada salvo às {ini.strftime('%H:%M')}.")
                st.rerun()
            except Exception as e:
                st.error(f"Não consegui salvar o início: {e}")


# ---------- pré-beneficiamento ----------
def tela_bags(etapa, chave, nomes, nominal=None, maizena=False, nota=""):
    st.subheader(etapa)
    if nota:
        st.caption(nota)
    st.caption("Assim que chegar, clique em **Iniciar observação** — isso já fica salvo na hora. "
               "Quando voltar (mesmo que seja no outro dia), abra esta tela de novo e encerre.")

    pend = pendentes(chave)
    if not pend.empty:
        linha = pend.sort_values("registrado_em", ascending=False).iloc[0]
        st.warning(f"Observação em andamento desde {linha['hora_ini']} "
                   f"({linha.get('tipo', '')}, {linha.get('aspereza', '')}). "
                   f"Preencha os dados abaixo para encerrar.")
        try:
            ini_time = datetime.strptime(str(linha["hora_ini"]), "%H:%M").time()
        except Exception:
            ini_time = time(0, 0)
        fim = campo_hora("Hora fim", chave + "_fimresume")
        dur = minutos(ini_time, fim)
        if dur <= 0:
            st.error("A hora fim precisa ser depois do início registrado.")
            return
        st.caption(f"Janela: {dur // 60}h{dur % 60:02d} ({dur} min)")

        with st.expander("Peso do bag"):
            peso = st.number_input("Peso do bag (kg)", 100, 2000, 900, 10, key=chave + "_peso")

        extra_key = chave + "_extras"
        if extra_key not in st.session_state:
            st.session_state[extra_key] = []
        c1, c2 = st.columns([3, 1])
        c1.markdown("**Bags por saída, no intervalo acima**")
        if c2.button("+ Adicionar saída", key=chave + "_addbtn", use_container_width=True):
            st.session_state[extra_key].append(f"Saída extra {len(st.session_state[extra_key]) + 1}")

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

        if st.button("Encerrar observação e salvar", key=chave + "_encerrar", type="primary"):
            liq = dur - par
            saidas_com_dado = [n for n in todas_saidas if kgs[n] > 0 and liq > 0]

            def monta(nome):
                v = kgs[nome] / 1000 / (liq / 60)
                pct = v / (nominal if nominal else 1.0) * 100
                d = {
                    "etapa": f"{etapa} — {nome}",
                    "hora_fim": fim.strftime("%H:%M"),
                    "duracao_min": dur,
                    "parado_min": par,
                    "motivo": mot,
                    "kg": round(kgs[nome], 1),
                    "vazao_tph": round(v, 3),
                    "pct_nominal": round(pct, 1),
                    "obs": obs,
                    "status_medicao": "Concluída",
                    "detalhe": f"{nome}: {kgs[nome] / peso:.2f} bag(s) de {peso}kg"
                }
                if maiz is not None:
                    d["maizena"] = maiz
                return d

            if not saidas_com_dado:
                atualizar_medicao(linha["id_medicao"],
                                  {"status_medicao": "Concluída", "hora_fim": fim.strftime("%H:%M")})
                st.info("Nenhuma saída tinha bags preenchidos; a observação foi encerrada sem gerar vazão.")
            else:
                primeira, *resto = saidas_com_dado
                full = atualizar_medicao(linha["id_medicao"], monta(primeira))
                if full:
                    _append_leitura(full)
                for nome in resto:
                    row = {
                        "data": linha.get("data", ""),
                        "hora_ini": linha["hora_ini"],
                        "tipo": linha.get("tipo", ""),
                        "aspereza": linha.get("aspereza", ""),
                        "lote": linha.get("lote", ""),
                        **monta(nome)
                    }
                    salvar(row)
                st.success("Observação encerrada e salva.")
            st.rerun()
        return

    data = st.date_input("Data", datetime.now(TZ).date(), key=chave + "_data")
    ini = campo_hora("Hora início", chave + "_ini")
    c1, c2 = st.columns(2)
    tipo = c1.selectbox("Tipo de feijão", TIPOS, key=chave + "_tipo")
    asp = c2.selectbox("Aspereza", ASPEREZA_OPCOES, key=chave + "_asp")
    lote = st.text_input("Lote (opcional)", key=chave + "_lote")
    if st.button("Iniciar observação (salva já)", key=chave + "_iniciar", type="primary"):
        row = {
            "id_medicao": novo_id(chave),
            "status_medicao": "Em andamento",
            "modo_medicao": "",
            "etapa": etapa,
            "data": str(data),
            "hora_ini": ini.strftime("%H:%M"),
            "hora_fim": "",
            "duracao_min": "",
            "parado_min": "",
            "kg": "",
            "vazao_tph": "",
            "pct_nominal": "",
            "motivo": "",
            "detalhe": "",
            "tipo": tipo,
            "aspereza": asp,
            "lote": lote,
            "maizena": "",
            "obs": "",
            "enfardadora": "",
            "silo_ini": "",
            "silo_fim": "",
            "registrado_em": datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
        }
        try:
            _append_medicoes(row)
            st.success(f"Início salvo às {ini.strftime('%H:%M')}. Volte aqui depois para encerrar.")
            st.rerun()
        except Exception as e:
            st.error(f"Não consegui salvar o início: {e}")


# ---------- beneficiamento via moega ----------
def tela_beneficiamento_moega():
    st.subheader("Beneficiamento (Moega)")
    st.caption("Acompanhamento do beneficiamento feito através da moega. "
               "Clique em **Iniciar observação** assim que abrir/iniciar a alimentação da moega. "
               "Ao esvaziar, informe a hora de término, quantidade de bags e pesos para calcular a vazão.")

    chave = "bene"
    nominal = 8.0  # Referência nominal de vazão (ton/h)

    pend = pendentes(chave)
    if not pend.empty:
        linha = pend.sort_values("registrado_em", ascending=False).iloc[0]
        st.warning(f"Observação em andamento desde {linha['hora_ini']} "
                   f"({linha.get('tipo', '')}, {linha.get('aspereza', '')}). "
                   f"Preencha os dados abaixo ao finalizar o lote na moega.")

        st.markdown("---")
        st.markdown("**1. Tipo de Feijão e Aspereza**")
        c_tipo, c_asp = st.columns(2)
        tipo_atual = linha.get("tipo", TIPOS[0])
        asp_atual = linha.get("aspereza", ASPEREZA_OPCOES[0])

        idx_tipo = TIPOS.index(tipo_
