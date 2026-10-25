import streamlit as st
import pandas as pd
import os
import datetime
from anthropic import Anthropic
import json
import base64
import shutil

# Product instellingen
st.set_page_config(page_title="Maritime AI Dropbox Module", page_icon="⚓", layout="centered")

EXCEL_FILE = "kasboek.xlsx"
FOTO_MAP = "fotos"
SLEUTEL_BESTAND = "sleutel.txt"
BACKUP_DOEL = "C:/Kasboek_Backups"

if not os.path.exists(FOTO_MAP):
    os.makedirs(FOTO_MAP)

# --- API KEY AUTO-SAVE LOGICA ---
if os.path.exists(SLEUTEL_BESTAND):
    with open(SLEUTEL_BESTAND, "r") as f:
        bewaarde_sleutel = f.read().strip()
else:
    bewaarde_sleutel = ""

st.title("⚓ Maritime AI-Dropbox Module")
st.caption("Proof of Concept Add-on voor Fleet Management Systemen (Marad / MXSuite)")

if not bewaarde_sleutel:
    st.subheader("🔑 1. Cloud AI Verbinding")
    api_key_input = st.text_input("Voer uw eenmalige API-sleutel in om de module te activeren:", type="password")
    if api_key_input:
        with open(SLEUTEL_BESTAND, "w") as f:
            f.write(api_key_input.strip())
        st.success("Verbinding succesvol geconfigureerd! Pagina wordt herladen...")
        st.rerun()
    api_key = api_key_input
else:
    api_key = bewaarde_sleutel

# --- EXCEL ARCHITECTUUR GENERATOR ---
def init_excel_sheets():
    if os.path.exists(EXCEL_FILE):
        return
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        raw_columns = [
            "ID", "Datum Invoer", "Datum Bon", "Bedrijf", "Transactienummer", 
            "Categorie", "Totaalbedrag (ANG)", "OB Bedrag (ANG)", "Fotoname", 
            "Literprijs (ANG)", "Aantal Liters", "Kilometerstand"
        ]
        pd.DataFrame(columns=raw_columns).to_excel(writer, sheet_name="RAW_DATA", index=False)
        
        for sheet_name in ["Boodschappen", "Non-Food", "Recreatie", "Auto_Motor"]:
            f_totaal = f'=SUMIF(RAW_DATA!F:F, "{sheet_name}", RAW_DATA!G:G)'
            f_ob = f'=SUMIF(RAW_DATA!F:F, "{sheet_name}", RAW_DATA!H:H)'
            if sheet_name == "Auto_Motor":
                df_sub = pd.DataFrame({
                    "Rapportage (Wal-Organisatie)": ["Jaartotaal YTD", "Lopend Maand Totaal", "Lopend Wekelijks Totaal"],
                    "Totaal Uitgegeven (ANG)": [f_totaal, f'=SUMIFS(RAW_DATA!G:G, RAW_DATA!F:F, "Auto_Motor")', ""],
                    "Totaal OB (ANG)": [f_ob, "", ""],
                    "Totaal KM Gereden": ["=IF(COUNT(RAW_DATA!L:L)>1, MAX(RAW_DATA!L:L)-MIN(RAW_DATA!L:L), 0)", "", ""],
                    "Gem. Verbruik (1 op X)": ["=IF(SUM(RAW_DATA!K:K)>0, (MAX(RAW_DATA!L:L)-MIN(RAW_DATA!L:L))/SUM(RAW_DATA!K:K), 0)", "", ""]
                })
            else:
                df_sub = pd.DataFrame({
                    "Rapportage (Wal-Organisatie)": ["Jaartotaal YTD", "Lopend Maand Totaal", "Lopend Wekelijks Totaal"],
                    "Totaal Uitgegeven (ANG)": [f_totaal, "", ""],
                    "Totaal OB (ANG)": [f_ob, "", ""]
                })
            df_sub.to_excel(writer, sheet_name=sheet_name, index=False)

# ============================= EINDE BLOK 1 =============================
# ============================= START BLOK 2 =============================

def verwerk_transactie_in_excel(cat, bedrijf, datum_bon, trans_nr, totaal, ob, ltr_prijs=0, liters=0, km_stand=None):
    init_excel_sheets()
    xls = pd.ExcelFile(EXCEL_FILE)
    df_raw = pd.read_excel(xls, "RAW_DATA")
    trans_id = len(df_raw) + 1
    tijdstempel = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    fotoname = f"{tijdstempel}_{cat}_{bedrijf}.jpg".replace(" ", "_").replace("/", "_")
    
    nieuwe_rij = {
        "ID": trans_id, "Datum Invoer": datetime.datetime.now().strftime('%Y-%m-%d'),
        "Datum Bon": datum_bon if datum_bon else datetime.datetime.now().strftime('%Y-%m-%d'),
        "Bedrijf": bedrijf if bedrijf else "Onbekend", "Transactienummer": trans_nr if trans_nr else "Geen",
        "Categorie": cat, "Totaalbedrag (ANG)": float(totaal) if totaal else 0.0, "OB Bedrag (ANG)": float(ob) if ob else 0.0,
        "Fotoname": fotoname, "Literprijs (ANG)": float(ltr_prijs) if ltr_prijs else 0.0,
        "Aantal Liters": float(liters) if liters else 0.0, "Kilometerstand": int(km_stand) if km_stand else None
    }
    df_raw = pd.concat([df_raw, pd.DataFrame([nieuwe_rij])], ignore_index=True)
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        df_raw.to_excel(writer, sheet_name="RAW_DATA", index=False)
        for sheet_name in ["Boodschappen", "Non-Food", "Recreatie", "Auto_Motor"]:
            f_totaal = f'=SUMIF(RAW_DATA!F:F, "{sheet_name}", RAW_DATA!G:G)'
            f_ob = f'=SUMIF(RAW_DATA!F:F, "{sheet_name}", RAW_DATA!H:H)'
            if sheet_name == "Auto_Motor":
                df_sub = pd.DataFrame({
                    "Rapportage (Wal-Organisatie)": ["Jaartotaal YTD", "Lopend Maand Totaal", "Lopend Wekelijks Totaal"],
                    "Totaal Uitgegeven (ANG)": [f_totaal, f'=SUMIFS(RAW_DATA!G:G, RAW_DATA!F:F, "Auto_Motor")', ""],
                    "Totaal OB (ANG)": [f_ob, "", ""],
                    "Totaal KM Gereden": ["=IF(COUNT(RAW_DATA!L:L)>1, MAX(RAW_DATA!L:L)-MIN(RAW_DATA!L:L), 0)", "", ""],
                    "Gem. Verbruik (1 op X)": ["=IF(SUM(RAW_DATA!K:K)>0, (MAX(RAW_DATA!L:L)-MIN(RAW_DATA!L:L))/SUM(RAW_DATA!K:K), 0)", "", ""]
                })
            else:
                df_sub = pd.DataFrame({
                    "Rapportage (Wal-Organisatie)": ["Jaartotaal YTD", "Lopend Maand Totaal", "Lopend Wekelijks Totaal"],
                    "Totaal Uitgegeven (ANG)": [f_totaal, "", ""],
                    "Totaal OB (ANG)": [f_ob, "", ""]
                })
            df_sub.to_excel(writer, sheet_name=sheet_name, index=False)
    return fotoname

def analyseer_fotos_met_ai(receipt_bytes, odometer_bytes=None):
    client = Anthropic(api_key=api_key)
    content_prompt = [
        {
            "type": "text",
            "text": "Analyseer de foto van deze bon/kilometerteller. Lever EXCLUSIEF een geldige JSON-string op met keys: 'bedrijf', 'datum_bon', 'transactienummer', 'literprijs', 'liters', 'totaalbedrag', 'ob_bedrag', 'kilometerstand'. Zet onbekende op null of 0. Geef GEEN markdown-codeblokken."
        },
        {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": receipt_bytes}}
    ]
    if odometer_bytes:
        content_prompt.append({"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": odometer_bytes}})
    message = client.messages.create(model="claude-sonnet-5-5", max_tokens=1000, messages=[{"role": "user", "content": content_prompt}])
    ai_tekst = "".join([blok.text for blok in message.content if hasattr(blok, "text")])
    return json.loads(ai_tekst.strip())

# --- USER INTERFACE TERMINAL ---
st.subheader("📱 2. Operationele Mobiele Invoer-Terminal")
col1, col2 = st.columns(2)
col3, col4 = st.columns(2)

if 'category' not in st.session_state: st.session_state.category = None
with col1:
    if st.button("🚗 Auto / Motor", use_container_width=True): st.session_state.category = "Auto_Motor"
with col2:
    if st.button("🛒 Boodschappen", use_container_width=True): st.session_state.category = "Boodschappen"
with col3:
    if st.button("📦 Non-Food", use_container_width=True): st.session_state.category = "Non-Food"
with col4:
    if st.button("🎉 Recreatie", use_container_width=True): st.session_state.category = "Recreatie"

if st.session_state.category:
    st.info(f"Actieve Modus: **{st.session_state.category}**")
    if not api_key:
        st.error("⚠️ Activeer eerst de Cloud AI Verbinding.")
    else:
        if st.session_state.category == "Auto_Motor":
            receipt_file = st.camera_input("Scan TANKBON")
            odometer_file = st.camera_input("Scan KILOMETERSTAND")
            if receipt_file and odometer_file:
                if st.button("Verwerk Brandstof & Rit-data", type="primary", use_container_width=True):
                    with st.spinner("AI analyseert data..."):
                        try:
                            rec_b64 = base64.b64encode(receipt_file.getvalue()).decode("utf-8")
                            odo_b64 = base64.b64encode(odometer_file.getvalue()).decode("utf-8")
                            ai = analyseer_fotos_met_ai(rec_b64, odo_b64)
                            fn = verwerk_transactie_in_excel("Auto_Motor", ai.get('bedrijf'), ai.get('datum_bon'), ai.get('transactienummer'), ai.get('totaalbedrag'), ai.get('ob_bedrag'), ai.get('literprijs'), ai.get('liters'), ai.get('kilometerstand'))
                            with open(os.path.join(FOTO_MAP, fn), "wb") as f: f.write(receipt_file.getvalue())
                            st.success("🎉 Geregistreerd en berekend!")
                        except Exception as e: st.error(f"Fout: {e}")
        else:
            receipt_file = st.camera_input("Scan DOCUMENT / BON")
            if receipt_file:
                if st.button(f"Analyseer & Boek {st.session_state.category}", type="primary", use_container_width=True):
                    with st.spinner("AI digitaliseert details..."):
                        try:
                            rec_b64 = base64.b64encode(receipt_file.getvalue()).decode("utf-8")
                            ai = analyseer_fotos_met_ai(rec_b64)
                            fn = verwerk_transactie_in_excel(st.session_state.category, ai.get('bedrijf'), ai.get('datum_bon'), ai.get('transactienummer'), ai.get('totaalbedrag'), ai.get('ob_bedrag'))
                            with open(os.path.join(FOTO_MAP, fn), "wb") as f: f.write(receipt_file.getvalue())
                            st.success(f"🎉 Geregistreerd onder {st.session_state.category}!")
                        except Exception as e: st.error(f"Fout: {e}")

# --- WAL-ORGANISATIE MODULE ---
st.divider()
st.subheader("🖥️ 3. FMS Controle-panel & Audit (Wal-organisatie)")
if os.path.exists(EXCEL_FILE):
    st.dataframe(pd.read_excel(EXCEL_FILE, sheet_name="RAW_DATA"), use_container_width=True)
    if st.button("💾 Genereer Gecentraliseerde Backup naar PC-schijf", type="secondary", use_container_width=True):
        try:
            tb = datetime.datetime.now().strftime('%Y%m%d')
            bp = os.path.join(BACKUP_DOEL, f"Backup_{tb}")
            if not os.path.exists(bp): os.makedirs(bp)
            shutil.copy(EXCEL_FILE, os.path.join(bp, EXCEL_FILE))
            if os.path.exists(os.path.join(bp, FOTO_MAP)): shutil.rmtree(os.path.join(bp, FOTO_MAP))
            shutil.copytree(FOTO_MAP, os.path.join(bp, FOTO_MAP))
            st.success(f"🔒 Back-up geëxporteerd: `{bp}/`")
        except Exception as e: st.error(f"Export mislukt: {e}")
else:
    st.info("Systeem staat paraat. Wacht op eerste operationele data-invoer.")
