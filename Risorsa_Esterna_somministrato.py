import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io

# ------------------------------------------------------------
# Caricamento configurazione da Excel caricato dall'utente
# ------------------------------------------------------------
def load_config_from_bytes(data: bytes):
    # Legge solo il foglio “Somministrato”
    cfg = pd.read_excel(io.BytesIO(data), sheet_name="Somministrato")
    # Separa le sezioni
    ou_df  = cfg[cfg["Section"] == "OU"][["Key/App", "Label/Gruppi/Value"]].rename(
                 columns={"Key/App": "key", "Label/Gruppi/Value": "label"})
    grp_df = cfg[cfg["Section"] == "InserimentoGruppi"][["Key/App", "Label/Gruppi/Value"]].rename(
                 columns={"Key/App": "app", "Label/Gruppi/Value": "gruppi"})
    def_df = cfg[cfg["Section"] == "Defaults"][["Key/App", "Label/Gruppi/Value"]].rename(
                 columns={"Key/App": "key", "Label/Gruppi/Value": "value"})

    ou_options = dict(zip(ou_df["key"], ou_df["label"]))
    gruppi     = dict(zip(grp_df["app"], grp_df["gruppi"]))
    defaults   = dict(zip(def_df["key"], def_df["value"]))
    return ou_options, gruppi, defaults

# ------------------------------------------------------------
# App 1.2: Risorsa Esterna - Somministrato/Stage
# ------------------------------------------------------------
st.set_page_config(page_title="1.2 Risorsa Esterna: Somministrato/Stage")
st.title("1.2 Risorsa Esterna: Somministrato/Stage")

config_file = st.file_uploader(
    "Carica il file di configurazione (config_corrected.xlsx)",
    type=["xlsx"],
    help="Deve contenere il foglio “Somministrato”"
)
if not config_file:
    st.warning("Per favore carica il file di configurazione per continuare.")
    st.stop()

ou_options, gruppi, defaults = load_config_from_bytes(config_file.read())

# ------------------------------------------------------------
# Utility functions
# ------------------------------------------------------------
def formatta_data(data: str) -> str:
    for sep in ["-", "/"]:
        try:
            g, m, a = map(int, data.split(sep))
            dt = datetime(a, m, g) + timedelta(days=1)
            return dt.strftime("%m/%d/%Y 00:00")
        except:
            continue
    return data

def genera_samaccountname(nome: str, cognome: str,
                          secondo_nome: str = "", secondo_cognome: str = "",
                          esterno: bool = False) -> str:
    n, sn = nome.strip().lower(), secondo_nome.strip().lower()
    c, sc = cognome.strip().lower(), secondo_cognome.strip().lower()
    suffix = ".ext" if esterno else ""
    limit  = 16 if esterno else 20
    cand   = f"{n}{sn}.{c}{sc}"
    if len(cand) <= limit:
        return cand + suffix
    cand = f"{n[:1]}{sn[:1]}.{c}{sc}"
    if len(cand) <= limit:
        return cand + suffix
    return (f"{n[:1]}{sn[:1]}.{c}")[:limit] + suffix

def build_full_name(cognome: str, secondo_cognome: str,
                    nome: str, secondo_nome: str,
                    esterno: bool = False) -> str:
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    full  = " ".join(parts)
    return full + (" (esterno)" if esterno else "")

HEADER = [
    "sAMAccountName","Creation","OU","Name","DisplayName","cn","GivenName","Surname",
    "employeeNumber","employeeID","department","Description","passwordNeverExpired",
    "ExpireDate","userprincipalname","mail","mobile","RimozioneGruppo","InserimentoGruppo",
    "disable","moveToOU","telephoneNumber","company"
]

# ------------------------------------------------------------
# Form di input nell’ordine richiesto
# ------------------------------------------------------------
cognome          = st.text_input("Cognome").strip().capitalize()
secondo_cognome  = st.text_input("Secondo Cognome").strip().capitalize()
nome             = st.text_input("Nome").strip().capitalize()
secondo_nome     = st.text_input("Secondo Nome").strip().capitalize()
codice_fiscale   = st.text_input("Codice Fiscale", "").strip()
department       = st.text_input("Sigla Divisione-Area", defaults.get("department_default", "")).strip()
numero_telefono  = st.text_input("Mobile", "").replace(" ", "")
description      = st.text_input("PC", "<PC>").strip()
expire_date      = st.text_input("Data di Fine (gg-mm-aaaa)", defaults.get("expire_default", "30-06-2025")).strip()

# ------------------------------------------------------------
# Valori fissi prelevati dalla configurazione
# ------------------------------------------------------------
ou_value           = defaults.get("ou_esterna_stage", "")
employee_id        = defaults.get("employee_id_default", "")
inserimento_gruppo = gruppi.get("esterna_stage", "")
telephone_number   = defaults.get("telephone_interna", "")
company            = defaults.get("company_interna", "")

# ------------------------------------------------------------
# Bottone di generazione CSV
# ------------------------------------------------------------
if st.button("Genera CSV"):
    sAM     = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, True)
    cn      = build_full_name(cognome, secondo_cognome, nome, secondo_nome, True)
    exp_fmt = formatta_data(expire_date)
    upn     = f"{sAM}@consip.it"
    mobile  = f"+39 {numero_telefono}" if numero_telefono else ""
    name    = cn
    display = cn
    given   = f"{nome} {secondo_nome}".strip()
    surn    = f"{cognome} {secondo_cognome}".strip()

    row = [
        sAM, "SI", ou_value, name, display, cn, given, surn,
        codice_fiscale, employee_id, department, description or "<PC>", "No", exp_fmt,
        upn, upn, mobile, "", inserimento_gruppo, "", "",
        telephone_number, company
    ]

    buf = io.StringIO()
    writer = csv.writer(buf, quoting=csv.QUOTE_MINIMAL)
    writer.writerow(HEADER)
    writer.writerow(row)
    buf.seek(0)

    st.dataframe(pd.DataFrame([row], columns=HEADER))
    st.download_button(
        label="📥 Scarica CSV Somministrato",
        data=buf.getvalue(),
        file_name=f"{cognome}_{nome[:1]}_stage.csv",
        mime="text/csv"
    )
    st.success(f"✅ File CSV generato per '{sAM}'")
