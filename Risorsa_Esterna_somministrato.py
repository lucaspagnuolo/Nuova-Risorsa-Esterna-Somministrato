import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io

# ------------------------------------------------------------
# Caricamento configurazione da Excel caricato dall'utente
# ------------------------------------------------------------
def load_config_from_bytes(data: bytes):
    cfg = pd.read_excel(io.BytesIO(data), sheet_name=None)
    ou_df = cfg.get("OU", pd.DataFrame(columns=["key", "label"]))
    ou_options = dict(zip(ou_df["key"], ou_df["label"]))
    grp_df = cfg.get("InserimentoGruppi", pd.DataFrame(columns=["app", "gruppi"]))
    gruppi = dict(zip(grp_df["app"], grp_df["gruppi"]))
    def_df = cfg.get("Defaults", pd.DataFrame(columns=["key", "value"]))
    defaults = dict(zip(def_df["key"], def_df["value"]))
    return ou_options, gruppi, defaults

# Uploader per config
st.title("1.2 Risorsa Esterna: Somministrato/Stage")
config_file = st.file_uploader(
    "Carica config.xlsx",
    type=["xlsx"], help="File con fogli OU, InserimentoGruppi e Defaults"
)
if not config_file:
    st.warning("Carica il file di configurazione per procedere.")
    st.stop()

ou_options, gruppi, defaults = load_config_from_bytes(config_file.read())

# Utility

def formatta_data(data: str) -> str:
    for sep in ["-", "/"]:
        try:
            g, m, a = map(int, data.split(sep))
            dt = datetime(a, m, g) + timedelta(days=1)
            return dt.strftime("%m/%d/%Y 00:00")
        except:
            continue
    return data


def genera_samaccountname(nome: str, cognome: str, secondo_nome: str = "", secondo_cognome: str = "", esterno: bool = False) -> str:
    n, sn = nome.strip().lower(), secondo_nome.strip().lower()
    c, sc = cognome.strip().lower(), secondo_cognome.strip().lower()
    suffix = ".ext" if esterno else ""
    limit = 16 if esterno else 20
    cand = f"{n}{sn}.{c}{sc}"
    if len(cand) <= limit: return cand + suffix
    cand = f"{(n[:1])}{(sn[:1])}.{c}{sc}"
    if len(cand) <= limit: return cand + suffix
    return (f"{n[:1]}{sn[:1]}.{c}")[:limit] + suffix


def build_full_name(cognome: str, secondo_cognome: str, nome: str, secondo_nome: str, esterno: bool = False) -> str:
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    full = " ".join(parts)
    return full + (" (esterno)" if esterno else "")

HEADER = [
    "sAMAccountName", "Creation", "OU", "Name", "DisplayName", "cn", "GivenName", "Surname",
    "employeeNumber", "employeeID", "department", "Description", "passwordNeverExpired",
    "ExpireDate", "userprincipalname", "mail", "mobile", "RimozioneGruppo", "InserimentoGruppo",
    "disable", "moveToOU", "telephoneNumber", "company"
]

# Sezione 1.2
nome            = st.text_input("Nome").strip().capitalize()
secondo_nome    = st.text_input("Secondo Nome").strip().capitalize()
cognome         = st.text_input("Cognome").strip().capitalize()
secondo_cognome = st.text_input("Secondo Cognome").strip().capitalize()
numero_telefono = st.text_input("Numero di Telefono", "").replace(" ", "")
description     = st.text_input("Description (lascia vuoto per <PC>)", "<PC>").strip()
codice_fiscale  = st.text_input("Codice Fiscale", "").strip()

# Valori fissi da config
ou_value         = ou_options.get("esterna_stage", defaults.get("ou_default", "Utenti esterni - Somministrati e Stage"))
expire_default   = defaults.get("expire_default", "30-06-2025")
employee_id      = defaults.get("employee_id_default", "")
department       = st.text_input("Dipartimento", defaults.get("department_default", "")).strip()
inserimento_gruppo = gruppi.get("esterna_stage", "")
telephone_number = defaults.get("telephone_default", "")
company          = defaults.get("company_default", "")

# Email flag input
email_flag = st.radio("Email necessaria?", ["Sì", "No"]) == "Sì"
if email_flag:
    try:
        custom_email = f"{cognome.lower()}{nome[0].lower()}@consip.it"
    except:
        st.error("Inserisci Nome e Cognome per email automatica.")
        custom_email = ""
else:
    custom_email = None

# Impostazione della data di scadenza
expire_date = st.text_input("Data di Fine (gg-mm-aaaa)", expire_default).strip()

if st.button("Genera CSV Esterna Stage"):
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, True)
    cn  = build_full_name(cognome, secondo_cognome, nome, secondo_nome, True)
    exp = formatta_data(expire_date)
    mail = custom_email if custom_email else f"{sAM}@consip.it"

    row = [
        sAM, "SI", ou_value, cn.replace(" (esterno)", ""), cn, cn,
        " ".join([nome, secondo_nome]).strip(),
        " ".join([cognome, secondo_cognome]).strip(),
        codice_fiscale, employee_id, department,
        description or "<PC>", "No", exp,
        f"{sAM}@consip.it", mail,
        f"+39 {numero_telefono}" if numero_telefono else "",
        "", inserimento_gruppo, "", "",
        telephone_number, company
    ]
    buf = io.StringIO()
    writer = csv.writer(buf, quoting=csv.QUOTE_MINIMAL)
    writer.writerow(HEADER)
    writer.writerow(row)
    buf.seek(0)

    st.dataframe(pd.DataFrame([row], columns=HEADER))
    st.download_button(
        "📥 Scarica CSV",
        buf.getvalue(),
        file_name=f"{cognome}_{nome[:1]}_stage.csv",
        mime="text/csv"
    )
    st.success(f"✅ File CSV generato per '{sAM}'")
