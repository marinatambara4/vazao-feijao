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
        return
