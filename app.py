import json
from datetime import datetime, time
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Vazão feijão", page_icon="🫘", layout="centered")
TZ = ZoneInfo("America/Sao_Paulo")
CSV_LOCAL = Path("medicoes.csv")
COLS = ["etapa", "data", "hora_ini", "hora_fim", "tipo", "aspereza", "lote",
        "duracao_min", "parado_min", "motivo", "kg", "vazao_tph", "pct_nominal",
        "maizena", "detalhe", "obs", "registrado_em"]

# ---------- armazenamento (Google Sheets ou CSV local) ----------
def usa_sheets():
    return "gcp_service_account" in st.secrets and "sheet_id" in st.secrets


@st.cache_resource
def get_ws():
    import gspread
    from google.oauth2.service_account import Credentials
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets"])
    sh = gspread.authorize(creds).open_by_key(st.secrets["sheet_id"])
    try:
        ws = sh.worksheet("medicoes")
    except Exception:
        ws = sh.add_worksheet("medicoes", rows=1000, cols=len(COLS))
    if not ws.row_values(1):
        ws.append_row(COLS)
    return ws


def salvar(row: dict):
    row["registrado_em"] = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
    vals = [row.get(c, "") for c in COLS]
    if usa_sheets():
        get_ws().append_row(vals, value_input_option="USER_ENTERED")
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


def cabecalho(p):
    data = st.date_input("Data", datetime.now(TZ).date(), key=p + "_data")
    ini = campo_hora("Hora início (bag fechando)", p + "_ini")
    fim = campo_hora("Hora fim (bag fechando)", p + "_fim")
    dur = minutos(ini, fim)
    if dur <= 0:
        st.error("A hora fim precisa ser depois da hora início.")
    else:
        st.caption(f"Janela: {dur // 60}h{dur % 60:02d} ({dur} min)")
    c1, c2 = st.columns(2)
    tipo = c1.selectbox("Tipo de feijão", ["Preto", "Carioca", "Vermelho"], key=p + "_tipo")
    asp = c2.radio("Aspereza", ["Liso", "Áspero"], horizontal=True, key=p + "_asp")
    lote = st.text_input("Lote (opcional)", key=p + "_lote")
    return dict(data=str(data), hora_ini=ini.strftime("%H:%M"), hora_fim=fim.strftime("%H:%M"),
                tipo=tipo, aspereza=asp, lote=lote, duracao_min=dur), dur


def paradas(p):
    c1, c2 = st.columns([1, 2])
    par = c1.number_input("Minutos parado", 0, 600, 0, key=p + "_par")
    mot = c2.text_input("Motivo da parada", key=p + "_mot")
    return par, mot


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


def confirmar(row):
    if row is None:
        return
    if st.button("Salvar medição", type="primary", use_container_width=True):
        try:
            salvar(row)
            st.success("Medição salva.")
        except Exception as e:
            st.error(f"Não salvou: {e}")


# ---------- telas de bags ----------
def tela_bags(etapa, chave, nominal, bocas_padrao, maizena=False, nota=""):
    st.subheader(etapa)
    if nota:
        st.caption(nota)
    base, dur = cabecalho(chave)
    with st.expander("Configurar bocas e peso do bag"):
        n = st.number_input("Número de bocas", 1, 6, bocas_padrao, key=chave + "_nb")
        peso = st.number_input("Peso do bag (kg)", 100, 2000, 900, 10, key=chave + "_peso")
    kg, det = 0.0, []
    st.markdown("**Bags completos por boca**")
    for i in range(int(n)):
        c1, c2, c3 = st.columns([2, 2, 2])
        nome = c1.text_input("Nome", f"Boca {i + 1}", key=f"{chave}_n{i}", label_visibility="collapsed")
        conta = c2.checkbox("Conta", value=(i == 0), key=f"{chave}_c{i}")
        bags = c3.number_input("Bags", 0, 100, 0, key=f"{chave}_b{i}", label_visibility="collapsed")
        det.append(f"{nome}:{bags}{'*' if conta else ''}")
        if conta:
            kg += bags * peso
    st.caption("Só as bocas marcadas em “Conta” entram na vazão (as demais ficam registradas).")
    par, mot = paradas(chave)
    extra = {"motivo": mot, "detalhe": f"bag {peso}kg | " + "; ".join(det)}
    if maizena:
        m = st.radio("Maizena usada?", ["Não", "Sim"], horizontal=True, key=chave + "_mz")
        extra["maizena"] = m
    extra["obs"] = st.text_area("Observações", key=chave + "_obs")
    confirmar(resultado({**base, "etapa": etapa}, kg, dur, par, nominal, extra))


def tela_embalagem():
    st.subheader("Embaladora + enfardadora")
    base, dur = cabecalho("emb")
    with st.expander("Peso por contagem do contador"):
        kgc = st.number_input("kg por unidade do contador", 0.1, 100.0, 1.0, 0.1, key="emb_kgc",
                              help="1 se o contador conta sacos de 1 kg; 30 se conta fardos.")
    kg, det, par_tot, mots = 0.0, [], 0, []
    for i in range(1, 5):
        with st.expander(f"Máquina {i}", expanded=(i == 1)):
            c1, c2 = st.columns(2)
            ini = c1.number_input("Contador início", 0, 100_000_000, 0, key=f"emb_ci{i}")
            fim = c2.number_input("Contador fim", 0, 100_000_000, 0, key=f"emb_cf{i}")
            c3, c4 = st.columns([1, 2])
            pm = c3.number_input("Min parada", 0, 600, 0, key=f"emb_pm{i}")
            mm = c4.text_input("Motivo", key=f"emb_mm{i}")
            k = max(fim - ini, 0) * kgc
            kg += k
            if k > 0 and dur - pm > 0:
                v = k / 1000 / ((dur - pm) / 60)
                st.caption(f"Máquina {i}: {v:.2f} ton/h (rodando {dur - pm} min)")
            det.append(f"M{i}:{int(k)}kg,par{pm}min")
            if pm and mm:
                mots.append(f"M{i}:{mm}")
    st.markdown("**Enfardadora e alimentação**")
    c1, c2 = st.columns(2)
    enf = c1.number_input("Enfardadora parada (min)", 0, 600, 0, key="emb_enf")
    falta = c2.number_input("Esperando feijão (min)", 0, 600, 0, key="emb_falta")
    mot_enf = st.text_input("Motivo (enfardadora)", key="emb_menf")
    pal = st.number_input("Paletes fechados na janela", 0, 500, 0, key="emb_pal",
                          help="Cada palete = 1.050 kg. Serve para conferir os contadores.")
    if pal:
        st.caption(f"Paletes: {pal * 1.05:.2f} t | Contadores: {kg / 1000:.2f} t")
    if enf:
        mots.append(f"Enfardadora:{mot_enf}")
    if falta:
        mots.append("Falta de feijão")
    extra = {"motivo": " | ".join(mots), "detalhe": "; ".join(det) + f"; paletes:{pal}; falta_feijao:{falta}min",
             "obs": st.text_area("Observações", key="emb_obs")}
    # vazão total sobre a janela inteira (paradas já refletem no contador)
    r = resultado({**base, "etapa": "Embaladora"}, kg, dur, 0, 6.0, extra)
    if r:
        r["parado_min"] = enf + falta
    confirmar(r)


# ---------- gráficos ----------
def tela_graficos():
    st.subheader("Comparação áspero × liso")
    df = carregar()
    if df.empty:
        st.info("Ainda não há medições salvas.")
        return
    df = df.dropna(subset=["vazao_tph"])
    tipos = ["Todos"] + sorted(df["tipo"].dropna().unique().tolist())
    sel = st.selectbox("Tipo de feijão", tipos)
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

    bene = df[(df["etapa"] == "Beneficiamento") & df["maizena"].isin(["Sim", "Não"])] if "maizena" in df else df.iloc[0:0]
    st.markdown("**4. Bene: com maizena × sem maizena**")
    if bene.empty:
        st.caption("Sem medições do bene com maizena registrada.")
    else:
        st.plotly_chart(px.box(bene, x="maizena", y="vazao_tph", color="aspereza", points="all",
                               color_discrete_map=cores), use_container_width=True)

    st.markdown("**5. Minutos parados por motivo**")
    p = df[(df["parado_min"] > 0) & (df["motivo"].astype(str) != "")]
    if not p.empty:
        pp = p.groupby("motivo")["parado_min"].sum().reset_index().sort_values("parado_min")
        st.plotly_chart(px.bar(pp, x="parado_min", y="motivo", orientation="h"), use_container_width=True)
    else:
        st.caption("Sem paradas registradas.")

    st.markdown("**6. Quantas medições por grupo**")
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
    tela_bags("Pré-beneficiamento", "pre", 14.5, 3,
              nota="Marque “Conta” só na boca do feijão tipo 1. Referência 14,5 ton/h (97% de 15).")
elif tela == "Bene":
    tela_bags("Beneficiamento", "bene", 6.0, 1, maizena=True)
elif tela == "Embaladora":
    tela_embalagem()
else:
    tela_graficos()
