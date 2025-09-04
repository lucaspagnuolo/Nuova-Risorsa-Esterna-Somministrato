import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io
import unicodedata
import zipfile

# ------------------------------------------------------------
# Caricamento configurazione da Excel caricato dall'utente
# ------------------------------------------------------------
def load_config_from_bytes(data: bytes):
    # Carica tutti i fogli del workbook
    cfg_sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, engine="openpyxl")

    # Estrai configurazione Somministrato
    sommin = cfg_sheets.get("Somministrato")
    grp_df = (
        sommin[sommin["Section"] == "InserimentoGruppi"]
        [["Key/App", "Label/Gruppi/Value"]]
        .rename(columns={"Key/App": "app", "Label/Gruppi/Value": "gruppi"})
    )
    gruppi = dict(zip(grp_df["app"], grp_df["gruppi"]))

    def_df = (
        sommin[sommin["Section"] == "Defaults"]
        [["Key/App", "Label/Gruppi/Value"]]
        .rename(columns={"Key/App": "key", "Label/Gruppi/Value": "value"})
    )
    defaults = dict(zip(def_df["key"], def_df["value"]))

    # Estrai opzioni manager dal foglio RA-RD
    rd_sheet = cfg_sheets.get("RA-RD")
    managers = {}
    if rd_sheet is not None:
        rd = rd_sheet.iloc[:, :2].dropna(how="all")
        rd.columns = ["label", "value"]
        managers = dict(zip(rd["label"], rd["value"]))

    # Estrai organigramma
    org_sheet = cfg_sheets.get("organigramma")
    organigramma = {}
    if org_sheet is not None:
        org = org_sheet.iloc[:, :2].dropna(how="all")
        org.columns = ["label", "value"]
        organigramma = dict(zip(org["label"], org["value"]))

    return gruppi, defaults, managers, organigramma

# ------------------------------------------------------------
# Utility functions
# ------------------------------------------------------------
def auto_quote(fields, quotechar='"', predicate=lambda s: ' ' in s):
    out = []
    for f in fields:
        s = str(f)
        if predicate(s):
            out.append(f'{quotechar}{s}{quotechar}')
        else:
            out.append(s)
    return out
    

def normalize_name(s: str) -> str:
    """Rimuove spazi, apostrofi e accenti, restituisce in minuscolo."""
    nfkd = unicodedata.normalize('NFKD', s)
    ascii_str = nfkd.encode('ASCII', 'ignore').decode()
    return ascii_str.replace(' ', '').replace("'", '').lower()


def formatta_data(data: str) -> str:
    for sep in ["-", "/"]:
        try:
            g, m, a = map(int, data.split(sep))
            dt = datetime(a, m, g) + timedelta(days=1)
            return dt.strftime("%m/%d/%Y 00:00")
        except:
            continue
    return data

# ------------------------------------------------------------
# Generazione SAMAccountName
# ------------------------------------------------------------
def genera_samaccountname(nome: str, cognome: str,
                          secondo_nome: str = "", secondo_cognome: str = "",
                          esterno: bool = True) -> str:
    n, sn = normalize_name(nome), normalize_name(secondo_nome)
    c, sc = normalize_name(cognome), normalize_name(secondo_cognome)
    suffix = ".ext" if esterno else ""
    limit  = 16 if esterno else 20

    cand1 = f"{n}{sn}.{c}{sc}"
    if len(cand1) <= limit:
        return cand1 + suffix
    cand2 = f"{n[:1]}{sn[:1]}.{c}{sc}"
    if len(cand2) <= limit:
        return cand2 + suffix
    base = f"{n[:1]}{sn[:1]}.{c}"
    return base[:limit] + suffix

# ------------------------------------------------------------
# Costruzione nome completo
# ------------------------------------------------------------
def build_full_name(cognome: str, secondo_cognome: str,
                    nome: str, secondo_nome: str,
                    esterno: bool = True) -> str:
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    return " ".join(parts) + (" (esterno)" if esterno else "")

# ------------------------------------------------------------
# Header CSV
# ------------------------------------------------------------
HEADER_USER = [
    "sAMAccountName","Creation","OU","Name","DisplayName","cn","GivenName","Surname",
    "employeeNumber","employeeID","department","Description","passwordNeverExpired",
    "ExpireDate","userprincipalname","mail","mobile","RimozioneGruppo","InserimentoGruppo",
    "disable","moveToOU","telephoneNumber","company"
]
HEADER_COMP = [
    "Computer","OU","add_mail","remove_mail","add_mobile","remove_mobile",
    "add_userprincipalname","remove_userprincipalname","disable","moveToOU"
]

# helper per creare buffer CSV
def make_csv_buffer(headers, row):
    buf = io.StringIO()
    w = csv.writer(buf, quoting=csv.QUOTE_NONE, escapechar="\\")
    quoted = auto_quote(row, quotechar='"', predicate=lambda s: ' ' in s)
    w.writerow(headers)
    w.writerow(quoted)
    buf.seek(0)
    return buf

# ------------------------------------------------------------
# Streamlit App
# ------------------------------------------------------------
st.set_page_config(page_title="1.2 Risorsa Esterna: Somministrato/Stage")
st.title("1.2 Risorsa Esterna: Somministrato/Stage")

config_file = st.file_uploader("Carica il file di configurazione (config.xlsx)", type=["xlsx"])
if not config_file:
    st.warning("Per favore carica il file di configurazione per continuare.")
    st.stop()

# Carica configurazione
gruppi, defaults, managers, organigramma = load_config_from_bytes(config_file.read())

# Valori di default
o365_groups = [
    defaults.get("grp_o365_standard","O365 Utenti Standard"),
    defaults.get("grp_o365_teams","O365 Teams Premium"),
    defaults.get("grp_o365_copilot","O365 Copilot Plus")]
grp_foorban        = defaults.get("grp_foorban","Foorban_Users")
grp_salesforce     = defaults.get("grp_salesforce", "")  # <-- lettura grp_salesforce
pillole            = defaults.get("pillole","Pillole formative Teams Premium")
ou_value           = defaults.get("ou_default","Somministrati e Stage")
expire_default     = defaults.get("expire_default","30-06-2025")
department_default = defaults.get("department_default","")
telephone_default  = defaults.get("telephone_interna","")
company            = defaults.get("company_interna","")
