import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io

# ------------------------------------------------------------
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
    # manteniamo la stringa grezza
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
# Streamlit App
# ------------------------------------------------------------
st.set_page_config(page_title='1.1 Nuova Risorsa Interna')
st.title('1.1 Nuova Risorsa Interna')

# Uploader
config_file = st.file_uploader("Carica config (config.xlsx)", type=['xlsx'],
    help="Foglio 'Risorsa Interna' con Section, Key/App, Label/Gruppi/Value")
if not config_file:
    st.warning("Carica il file di configurazione.")
    st.stop()

ou_options,gruppi,defaults = load_config_from_bytes(config_file.read())

# Preleva valori Defaults
dl_standard = defaults.get('dl_standard','').split(';')
dl_vip      = defaults.get('dl_vip','').split(';')
o365_groups = [
    defaults.get('grp_o365_standard','O365 Utenti Standard'),
    defaults.get('grp_o365_teams','O365 Teams Premium'),
    defaults.get('grp_o365_copilot','O365 Copilot Plus')
]
grp_foorban = defaults.get('grp_foorban','Foorban_Users')
pillole     = defaults.get('pillole','Pillole formative Teams Premium')

# Input anagrafici
st.subheader('Modulo Inserimento Nuova Risorsa Interna')
employee_id = st.text_input('Matricola',defaults.get('employee_id_default','')).strip()
cognome     = st.text_input('Cognome').strip().capitalize()
secondo_cognome = st.text_input('Secondo Cognome').strip().capitalize()
nome        = st.text_input('Nome').strip().capitalize()
secondo_nome= st.text_input('Secondo Nome').strip().capitalize()
codice_fiscale = st.text_input('Codice Fiscale').strip()
department  = st.text_input('Sigla Divisione-Area',defaults.get('department_default','')).strip()
numero_telefono = st.text_input('Mobile (+39 già inserito)').replace(' ','')
description = st.text_input('PC (lascia vuoto per <PC>)','<PC>').strip()

# Resident flag
resident = st.checkbox('È Resident?')
numero_fisso = ''
if resident:
    numero_fisso = st.text_input('Numero fisso Resident (+39 già inserito)').strip()
telephone_default = defaults.get('telephone_interna','')
telephone_number = f'+39 {numero_fisso}' if resident and numero_fisso else telephone_default

# Tipologia Utente
ou_keys = list(ou_options.keys())
ou_vals = list(ou_options.values())
def_o = defaults.get('ou_default',ou_vals[0] if ou_vals else '')
label_ou = st.selectbox('Tipologia Utente',ou_vals,index=ou_vals.index(def_o))
selected_key = ou_keys[ou_vals.index(label_ou)]
ou_value = ou_options[selected_key]

inserimento_gruppo = gruppi.get('interna','')
company = defaults.get('company_interna','')

# Operatività e SM
st.subheader('Configurazione Data Operatività e Profilazione SM')
data_operativa = st.text_input('In che giorno prende operatività? (gg/mm/aaaa)').strip()
profilazione = st.checkbox('Deve essere profilato su qualche SM?')
sm_lines = []
if profilazione:
    sm_lines = st.text_area('SM su quali va profilato','',placeholder='Inserisci una SM per riga').splitlines()

# Selezione DL
dl_list = dl_standard if selected_key=='utenti_standard' else dl_vip if selected_key=='utenti_vip' else []

# Anteprima
if st.button('Template per Posta Elettronica'):
    sAM = f"{nome.lower()}.{cognome.lower()}"
    cn = f"{cognome} {nome}"
    groups_md = "\n".join(f"- {g}" for g in o365_groups)
    table = f"""
| Campo             | Valore                                     |
|-------------------|--------------------------------------------|
| Tipo Utenza       | Remota                                     |
| Utenza            | {sAM}                                      |
| Alias             | {sAM}                                      |
| Display name      | {cn}                                       |
| Common name       | {cn}                                       |
| e-mail            | {sAM}@consip.it                            |
| e-mail secondaria | {sAM}@consipspa.mail.onmicrosoft.com       |
| cell              | +39 {numero_telefono}                      |
"""
    st.markdown(f"Ciao.  \nRichiedo la definizione di una casella come sottoindicato.")
    st.markdown(table)
    st.markdown(f"Inviare batch di notifica migrazione mail a: imac@consip.it  \n"+
                f"Aggiungere utenza di dominio ai gruppi:\n{groups_md}")
    if dl_list:
        st.markdown(f"Il giorno **{data_operativa}** occorre inserire la casella nelle DL:")
        for dl in dl_list:
            st.markdown(f"- {dl}")
    if profilazione:
        st.markdown('Profilare su SM:')
        for sm in sm_lines:
            st.markdown(f"- {sm}")
    st.markdown(f"Aggiungere utenza al:\n- gruppo Azure: {grp_foorban}\n- canale {pillole}")
    st.markdown('Grazie  \nSaluti')

# CSV
if st.button('Genera CSV Interna'):
    sAM = f"{nome.lower()}.{cognome.lower()}"
    cn = f"{cognome} {nome}"
    mobile = f"+39 {numero_telefono}" if numero_telefono else ''
    row = [sAM,'SI',ou_value,cn,cn,cn,nome,nome,
           codice_fiscale,employee_id,department,description,'No','',
           f"{sAM}@consip.it",f"{sAM}@consip.it",mobile,'',
           inserimento_gruppo,'','',telephone_number,company]
    buf = io.StringIO();writer=csv.writer(buf,quoting=csv.QUOTE_NONE,escapechar='\\')
    writer.writerow(HEADER);writer.writerow(row);buf.seek(0)
    st.dataframe(pd.DataFrame([row],columns=HEADER))
    st.download_button('📥 Scarica CSV',buf.getvalue(),f"{cognome}_{nome[:1]}_interno.csv","text/csv")
    st.success(f"✅ CSV generato per {sAM}")
