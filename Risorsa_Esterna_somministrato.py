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
    st.markdown(f"Aggiungere utenza al gruppo Azure: \n- {grp_foorban}\n- {grp_salesforce}\n- canale {pillole}")
    if profilazione_flag:
        st.markdown("Profilare su SM:")
        for sm in sm_lines:
            if sm.strip(): st.markdown(f"- {sm}")
    st.markdown("Grazie  \nSaluti")

# Generazione CSV Utente + Computer + Profilazione
if st.button("Genera CSV Somministrato"):
    sAM    = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, True)
    cn     = build_full_name(cognome, secondo_cognome, nome, secondo_nome, True)
    exp_fmt= formatta_data(expire_date)
    upn    = f"{sAM}@consip.it"
    mobile = f"+39 {numero_telefono}" if numero_telefono else ""
    given  = f"{nome} {secondo_nome}".strip()
    surn   = f"{cognome} {secondo_cognome}".strip()

    # Costruisci basename normalizzato (visibile nei file)
    norm_cognome = normalize_name(cognome)
    norm_secondo = normalize_name(secondo_cognome) if secondo_cognome else ''
    name_parts = [norm_cognome] + ([norm_secondo] if norm_secondo else []) + [nome[:1].lower()]
    basename = "_".join(name_parts)
    name_parts = [cognome] + ([secondo_cognome] if secondo_cognome else []) + [nome[:1]]
    basename = "_".join(name_parts)

    # Righe CSV Utente: InserimentoGruppo lasciato volutamente vuoto come richiesto
    row_user = [
        sAM, "SI", ou_value, cn, cn, cn, given, surn,
        codice_fiscale, employee_id, department, description or "<PC>",
        "No", exp_fmt, upn, upn, mobile,
        "", "", "", "", telephone_number, company
    ]
    row_comp = [
        description or "", "", f"{sAM}@consip.it", "", f'"{mobile}"', "",
        f'"{cn}"', "", "", ""
    ]

    # Profilazione: costruisco lista gruppi unendo o365_groups e inserimento_gruppo (filtrando vuoti)
    profile_groups_list = []
    for g in o365_groups:
        if g and str(g).strip():
            token = str(g).strip()
            # Piccola correzione automatica: se per errore manca la "O" iniziale (es. "365 Utenti Standard")
            if token.startswith("365 "):
                token = "O" + token
            profile_groups_list.append(token)
    if inserimento_gruppo and str(inserimento_gruppo).strip():
        ig = str(inserimento_gruppo).strip()
        profile_groups_list.append(ig)

    # Join senza spazi dopo il punto e virgola (produrra: "A;B;C;d;...")
    profile_groups = ";".join(profile_groups_list)

    # Costruisco riga Profilazione con stesso header di HEADER_USER, ma valorizzando solo sAMAccountName e InserimentoGruppo
    profile_row = [""] * len(HEADER_USER)
    profile_row[0] = sAM
    try:
        idx_inserimento = HEADER_USER.index("InserimentoGruppo")
    except ValueError:
        idx_inserimento = 18
    profile_row[idx_inserimento] = profile_groups

    # --- Messaggi personalizzati per utente, computer e profilazione ---
    msg_utente = (
        "Salve.\n"
        "Vi richiediamo la definizione della utenza nell\u2019AD Consip come dettagliato nei file:\n"
        f"\\\\srv_dati\\AreaCondivisa\\DEPSI\\IC\\Utenze\\Esterni\\{basename}_utente.csv\n"
        "Restiamo in attesa di un vostro riscontro ad attivit\u00e0 completata.\n"
        "Saluti"
    )

    msg_computer = (
        "Salve.\n"
        "Si richiede modifiche come da file:\n"
        f"\\\\srv_dati\\AreaCondivisa\\DEPSI\\IC\\PC\\{basename}_computer.csv\n"
        "Restiamo in attesa di un vostro riscontro ad attivit\u00e0 completata.\n"
        "Saluti"
    )

    msg_profilazione = (
        "Salve.\n"
        "Si richiede modifiche come da file:\n"
        f"\\\\srv_dati\\AreaCondivisa\\DEPSI\\IC\\Profilazioni\\{basename}_profilazione.csv\n"
        "Restiamo in attesa di un vostro riscontro ad attivit\u00e0 completata.\n"
        "Saluti"
    )

    # --- mostra a video i messaggi sintetici e le anteprime ---
    st.subheader(f"Nuova Utenza AD [{cognome} (esterno)]")
    st.text(msg_utente)
    st.subheader("Anteprima CSV Utente")
    st.dataframe(pd.DataFrame([row_user], columns=HEADER_USER))

    # Download CSV Utente (sotto l'anteprima utente)
    buf_user = make_csv_buffer(HEADER_USER, row_user)
    st.download_button(
        "📥 Scarica CSV Utente",
        data=buf_user.getvalue(),
        file_name=f"{basename}_utente.csv",
        mime="text/csv"
    )

    st.subheader(f"Modifica AD [{cognome} (esterno)]")
    st.text(msg_computer)
    st.subheader("Anteprima CSV Computer")
    st.dataframe(pd.DataFrame([row_comp], columns=HEADER_COMP))

    # Download CSV Computer (sotto l'anteprima computer)
    buf_comp = make_csv_buffer(HEADER_COMP, row_comp)
    st.download_button(
        "📥 Scarica CSV Computer",
        data=buf_comp.getvalue(),
        file_name=f"{basename}_computer.csv",
        mime="text/csv"
    )

    st.subheader(f"Modifica AD Profilazione [{cognome} (esterno)]")
    st.text(msg_profilazione)
    st.subheader("Anteprima CSV Profilazione")
    st.dataframe(pd.DataFrame([profile_row], columns=HEADER_USER))

    # Download CSV Profilazione (sotto l'anteprima profilazione)
    buf_prof = make_csv_buffer(HEADER_USER, profile_row)
    st.download_button(
        "📥 Scarica CSV Profilazione",
        data=buf_prof.getvalue(),
        file_name=f"{basename}_profilazione.csv",
        mime="text/csv"
    )

    # -------------------------------------------
    # Preparo l'anteprima template in formato Markdown (da inserire SOLO nello ZIP)
    # -------------------------------------------
    table_md = (
        "| Campo             | Valore                                     |\n"
        "|-------------------|--------------------------------------------|\n"
        f"| Tipo Utenza       | Remota                                     |\n"
        f"| Utenza            | {sAM}                                       |\n"
        f"| Alias             | {sAM}                                       |\n"
        f"| Display name      | {cn}                                        |\n"
        f"| Common name       | {cn}                                        |\n"
        f"| e-mail            | {upn}                                       |\n"
        f"| e-mail secondaria | {upn}                                       |\n"
    )

    template_preview_lines = []
    template_preview_lines.append("Richiesta definizione casella - anteprima template\n")
    template_preview_lines.append(table_md)
    if profile_groups_list:
        template_preview_lines.append(f"\nIl giorno {expire_date} occorre inserire la casella nelle DL:\n")
        for dl in profile_groups_list:
            template_preview_lines.append(f"- {dl}\n")
    if profilazione_flag:
        template_preview_lines.append("\nProfilare su SM:\n")
        for sm in sm_lines:
            template_preview_lines.append(f"- {sm}\n")

    if grp_foorban:
        template_preview_lines.append(f"\nAggiungere utenza al gruppo Azure:\n- {grp_foorban}\n")
    if grp_salesforce:
        template_preview_lines.append(f"- {grp_salesforce}\n")
    if pillole:
        template_preview_lines.append(f"- canale {pillole}\n")

    # unisco tutto in una stringa markdown
    template_preview_md = "\n".join(template_preview_lines)

    # -------------------------------------------
    # Creo lo ZIP con i 3 CSV + file di anteprima template + 3 messaggi .txt
    # (i file "msg_*.txt" e "template_preview.md" saranno presenti SOLO nello ZIP)
    # -------------------------------------------
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        # CSV
        zipf.writestr(f"{basename}_utente.csv", buf_user.getvalue())
        zipf.writestr(f"{basename}_computer.csv", buf_comp.getvalue())
        zipf.writestr(f"{basename}_profilazione.csv", buf_prof.getvalue())
        # anteprima template (markdown)
        zipf.writestr(f"{basename}_template_preview.md", template_preview_md)
        # messaggi (solo testo)
        zipf.writestr(f"{basename}_msg_utente.txt", msg_utente)
        zipf.writestr(f"{basename}_msg_computer.txt", msg_computer)
        zipf.writestr(f"{basename}_msg_profilazione.txt", msg_profilazione)

    zip_buffer.seek(0)

    # pulsante per scaricare l'unico ZIP (contenente anche template + messaggi)
    st.download_button(
        "📦 Scarica Tutti i CSV (ZIP) + anteprima e messaggi",
        data=zip_buffer.getvalue(),
        file_name=f"{basename}_csv_bundle.zip",
        mime="application/zip"
    )

    st.success(f"✅ CSV generati per '{sAM}'")
