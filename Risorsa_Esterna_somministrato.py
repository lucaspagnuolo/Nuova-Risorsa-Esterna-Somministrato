import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io
import unicodedata

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
pillole            = defaults.get("pillole","Pillole formative Teams Premium")
ou_value           = defaults.get("ou_default","Somministrati e Stage")
expire_default     = defaults.get("expire_default","30-06-2025")
department_default = defaults.get("department_default","")
telephone_default  = defaults.get("telephone_interna","")
company            = defaults.get("company_interna","")

# Modulo di input
st.subheader("Modulo Inserimento Risorsa Esterna: Somministrato/Stage")
cognome          = st.text_input("Cognome").strip().capitalize()
secondo_cognome  = st.text_input("Secondo Cognome").strip().capitalize()
nome             = st.text_input("Nome").strip().capitalize()
secondo_nome     = st.text_input("Secondo Nome").strip().capitalize()
codice_fiscale   = st.text_input("Codice Fiscale","" ).strip()

# Dropdown per Sigla Divisione-Area da organigramma
if organigramma:
    dept_label = st.selectbox("Sigla Divisione-Area", options=["-- Seleziona --"] + list(organigramma.keys()))
    department = organigramma.get(dept_label, "") if dept_label and dept_label != "-- Seleziona --" else department_default
else:
    department = st.text_input("Sigla Divisione-Area", department_default).strip()

numero_telefono  = st.text_input("Mobile","" ).replace(" ","")
description      = st.text_input("PC (lascia vuoto per <PC>)","<PC>" ).strip()
expire_date      = st.text_input("Data di Fine (gg-mm-aaaa)", expire_default).strip()
profilazione_flag= st.checkbox("Deve essere profilato su qualche SM?")
sm_lines         = st.text_area("SM su quali va profilato","" ).splitlines() if profilazione_flag else []
employee_id      = ""
inserimento_gruppo= gruppi.get("esterna_stage", "")
telephone_number = telephone_default

# Dropdown Manager
if managers:
    manager_label = st.selectbox("Manager", options=["-- Seleziona --"] + list(managers.keys()))
    manager = managers.get(manager_label, "") if manager_label and manager_label != "-- Seleziona --" else ""
else:
    manager = st.text_input("Manager").strip()

# Anteprima Messaggio (Template)
if st.button("Template per Posta Elettronica"):
    sAM    = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, True)
    cn     = build_full_name(cognome, secondo_cognome, nome, secondo_nome, True)
    exp_fmt= formatta_data(expire_date)
    upn    = f"{sAM}@consip.it"
    mobile = f"+39 {numero_telefono}" if numero_telefono else ""

    st.markdown("""
Ciao.
Richiedo la definizione di una casella come sottoindicato.
""")
    table_md = f"""
| Campo             | Valore                                     |
|-------------------|--------------------------------------------|
| Tipo Utenza       | Remota                                     |
| Utenza            | {sAM}                                       |
| Alias             | {sAM}                                       |
| Display name      | {cn}                                        |
| Common name       | {cn}                                        |
| Manager           | {manager}                                   |
| e-mail            | {upn}                                       |
| e-mail secondaria | {upn}                                       |
| Codice Fiscale    | {codice_fiscale}                            |
| Data Fine         | {exp_fmt}                                   |
"""
    st.markdown(table_md)
    st.markdown("** il campo \"Data fine\" deve essere inserito in \"Data Assunzione\" **")
    st.markdown("Aggiungere utenza di dominio ai gruppi:\n" + "\n".join(f"- {g}" for g in o365_groups))
    st.markdown(f"Aggiungere utenza al:\n- gruppo Azure: {grp_foorban}\n- canale {pillole}")
    if profilazione_flag:
        st.markdown("Profilare su SM:")
        for sm in sm_lines:
            if sm.strip(): st.markdown(f"- {sm}")
    st.markdown("Grazie  \nSaluti")

# Generazione CSV Utente + Computer
if st.button("Genera CSV Somministrato"):
    sAM    = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, True)
    cn     = build_full_name(cognome, secondo_cognome, nome, secondo_nome, True)
    exp_fmt= formatta_data(expire_date)
    upn    = f"{sAM}@consip.it"
    mobile = f"+39 {numero_telefono}" if numero_telefono else ""
    given  = f"{nome} {secondo_nome}".strip()
    surn   = f"{cognome} {secondo_cognome}".strip()

    # Costruisci basename normalizzato
    norm_cognome = normalize_name(cognome)
    norm_secondo = normalize_name(secondo_cognome) if secondo_cognome else ''
    name_parts = [norm_cognome] + ([norm_secondo] if norm_secondo else []) + [nome[:1].lower()]
    basename = "_".join(name_parts)
    name_parts = [cognome] + ([secondo_cognome] if secondo_cognome else []) + [nome[:1]]
    basename = "_".join(name_parts)

    # Righe CSV
    row_user = [
        sAM, "SI", ou_value, cn, cn, cn, given, surn,
        codice_fiscale, employee_id, department, description or "<PC>",
        "No", exp_fmt, upn, upn, mobile,
        "", inserimento_gruppo, "", "", telephone_number, company
    ]
    row_comp = [
        description or "", "", f"{sAM}@consip.it", "", f"\"{mobile}\"", "",
        f"\"{cn}\"", "", "", ""
    ]

    # Preview messaggio
    st.markdown(f"""
Ciao.  
Si richiede modifiche come da file:  
- {basename}_computer.csv  (oggetti di tipo computer)  
- {basename}_utente.csv  (oggetti di tipo utenze)  

Archiviati al percorso:  
\\\\\\srv_dati.consip.tesoro.it\AreaCondivisa\DEPSI\IC\AD_Modifiche  
Grazie
"""
    )
    # Anteprime
    st.subheader("Anteprima CSV Utente")
    st.dataframe(pd.DataFrame([row_user], columns=HEADER_USER))
    st.subheader("Anteprima CSV Computer")
    st.dataframe(pd.DataFrame([row_comp], columns=HEADER_COMP))

    # Download
    buf_user = io.StringIO()
    w1 = csv.writer(buf_user, quoting=csv.QUOTE_NONE, escapechar="\\")
    quoted_row_user = auto_quote(row_user, quotechar='"', predicate=lambda s: ' ' in s)
    w1.writerow(HEADER_USER)
    w1.writerow(quoted_row_user)
    buf_user.seek(0)

    buf_comp = io.StringIO()
    w2 = csv.writer(buf_comp, quoting=csv.QUOTE_NONE, escapechar="\\")
    quoted_row_comp = auto_quote(row_comp, quotechar='"', predicate=lambda s: ' ' in s)
    w2.writerow(HEADER_COMP)
    w2.writerow(quoted_row_comp)
    buf_comp.seek(0)

    st.download_button(
        "📥 Scarica CSV Utente",
        data=buf_user.getvalue(),
        file_name=f"{basename}_utente.csv",
        mime="text/csv"
    )
    st.download_button(
        "📥 Scarica CSV Computer",
        data=buf_comp.getvalue(),
        file_name=f"{basename}_computer.csv",
        mime="text/csv"
    )
    st.success(f"✅ CSV generati per '{sAM}'")
