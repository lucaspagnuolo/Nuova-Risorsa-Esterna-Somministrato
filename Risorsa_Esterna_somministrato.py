import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io

# ------------------------------------------------------------
# Caricamento configurazione da Excel caricato dall'utente
# ------------------------------------------------------------
def load_config_from_bytes(data: bytes):
    cfg = pd.read_excel(io.BytesIO(data), sheet_name="Somministrato")
    # InserimentoGruppi (stringhe con ; separate)
    grp_df = (
        cfg[cfg["Section"] == "InserimentoGruppi"]
        [["Key/App", "Label/Gruppi/Value"]]
        .rename(columns={"Key/App": "app", "Label/Gruppi/Value": "gruppi"})
    )
    gruppi = dict(zip(grp_df["app"], grp_df["gruppi"]))

    # Defaults
    def_df = (
        cfg[cfg["Section"] == "Defaults"]
        [["Key/App", "Label/Gruppi/Value"]]
        .rename(columns={"Key/App": "key", "Label/Gruppi/Value": "value"})
    )
    defaults = dict(zip(def_df["key"], def_df["value"]))
    return gruppi, defaults

# ------------------------------------------------------------
# App 1.2: Risorsa Esterna - Somministrato/Stage
# ------------------------------------------------------------
st.set_page_config(page_title="1.2 Risorsa Esterna: Somministrato/Stage")
st.title("1.2 Risorsa Esterna: Somministrato/Stage")

config_file = st.file_uploader(
    "Carica il file di configurazione (config.xlsx)",
    type=["xlsx"],
    help="Deve contenere il foglio 'Somministrato' con colonne Section, Key/App, Label/Gruppi/Value"
)
if not config_file:
    st.warning("Per favore carica il file di configurazione per continuare.")
    st.stop()

gruppi, defaults = load_config_from_bytes(config_file.read())

# ------------------------------------------------------------
# Preleva valori Defaults
# ------------------------------------------------------------
o365_groups = [
    defaults.get("grp_o365_standard", "O365 Utenti Standard"),
    defaults.get("grp_o365_teams",    "O365 Teams Premium"),
    defaults.get("grp_o365_copilot",  "O365 Copilot Plus")
]
ou_value          = defaults.get("ou_default", "Somministrati e Stage")
expire_default    = defaults.get("expire_default", "30-06-2025")
department_default= defaults.get("department_default", "")
telephone_default = defaults.get("telephone_interna", "")
company           = defaults.get("company_interna", "")

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
                           esterno: bool = True) -> str:
    n, sn = nome.strip().lower(), secondo_nome.strip().lower()
    c, sc = cognome.strip().lower(), secondo_cognome.strip().lower()
    suffix = ".ext"
    cand = f"{n}{sn}.{c}{sc}"
    return (cand + suffix)[:16+len(suffix)]

def build_full_name(cognome: str, secondo_cognome: str,
                    nome: str, secondo_nome: str,
                    esterno: bool = True) -> str:
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    return " ".join(parts) + " (esterno)"

HEADER = [
    "sAMAccountName","Creation","OU","Name","DisplayName","cn","GivenName","Surname",
    "employeeNumber","employeeID","department","Description","passwordNeverExpired",
    "ExpireDate","userprincipalname","mail","mobile","RimozioneGruppo","InserimentoGruppo",
    "disable","moveToOU","telephoneNumber","company"
]

# ------------------------------------------------------------
# Form di input
# ------------------------------------------------------------
st.subheader("Modulo Inserimento Risorsa Esterna: Somministrato/Stage")

cognome         = st.text_input("Cognome").strip().capitalize()
secondo_cognome = st.text_input("Secondo Cognome").strip().capitalize()
nome            = st.text_input("Nome").strip().capitalize()
secondo_nome    = st.text_input("Secondo Nome").strip().capitalize()
codice_fiscale  = st.text_input("Codice Fiscale", "").strip()
department      = st.text_input("Sigla Divisione-Area", department_default).strip()
numero_telefono = st.text_input("Mobile", "").replace(" ", "")
description     = st.text_input("PC (lascia vuoto per <PC>)", "").strip()
expire_date     = st.text_input("Data di Fine (gg-mm-aaaa)", expire_default).strip()

# Profilazione SM
profilazione_flag = st.checkbox("Deve essere profilato su qualche SM?")
sm_lines = []
if profilazione_flag:
    sm_lines = st.text_area(
        "SM su quali va profilato", "", placeholder="Inserisci una SM per riga"
    ).splitlines()

# Dati fissi
employee_id        = ""  # sempre vuoto
inserimento_gruppo = gruppi.get("esterna_stage", "")
telephone_number   = telephone_default
company            = company

# ------------------------------------------------------------
# Anteprima Messaggio
# ------------------------------------------------------------
if st.button("Anteprima Messaggio"):
    sAM     = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, True)
    cn      = build_full_name(cognome, secondo_cognome, nome, secondo_nome, True)
    exp_fmt = formatta_data(expire_date)
    upn     = f"{sAM}@consip.it"
    mobile  = f"+39 {numero_telefono}" if numero_telefono else ""

    table_md = f"""
| Campo             | Valore                                     |
|-------------------|--------------------------------------------|
| Tipo Utenza       | Remota                                     |
| Utenza            | {sAM}                                      |
| Alias             | {sAM}                                      |
| Display name      | {cn}                                       |
| Common name       | {cn}                                       |
| e-mail            | {upn}                                      |
| e-mail secondaria | {upn}                                      |
"""
    st.markdown("Ciao.\nRichiedo la definizione di una casella come sottoindicato.")
    st.markdown(table_md)

    groups_md = "\n".join(f"- {g}" for g in o365_groups)
    st.markdown(
        f"Inviare batch di notifica migrazione mail a: imac@consip.it  \n"
        f"Aggiungere utenza di dominio ai gruppi:\n{groups_md}"
    )

    if profilazione_flag:
        st.markdown("Profilare su SM:")
        for sm in sm_lines:
            if sm.strip(): st.markdown(f"- {sm}")

    st.markdown("Grazie  \nSaluti")

# ------------------------------------------------------------
# Generazione CSV
# ------------------------------------------------------------
if st.button("Genera CSV Somministrato"):
    sAM     = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, True)
    cn      = build_full_name(cognome, secondo_cognome, nome, secondo_nome, True)
    exp_fmt = formatta_data(expire_date)
    upn     = f"{sAM}@consip.it"
    mobile  = f"+39 {numero_telefono}" if numero_telefono else ""
    given   = f"{nome} {secondo_nome}".strip()
    surn    = f"{cognome} {secondo_cognome}".strip()

    row = [
        sAM, "SI", ou_value,
        cn, cn, cn, given, surn,
        codice_fiscale, employee_id, department, description or "<PC>", "No", exp_fmt,
        upn, upn, mobile, "", inserimento_gruppo, "", "",
        telephone_number, company
    ]

    buf = io.StringIO()
    writer = csv.writer(buf, quoting=csv.QUOTE_NONE, escapechar="\\")

    # Quote i campi necessari
    for i in (2,3,4,5): row[i] = f"\"{row[i]}\""
    row[13] = f"\"{row[13]}\""
    row[16] = f"\"{row[16]}\""

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
