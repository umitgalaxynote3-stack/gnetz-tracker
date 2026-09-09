import subprocess
import sys
import os

# 1. AUTOMATISCHE INSTALLATION: Installiert das Excel-Paket im Hintergrund, falls es fehlt
try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])

import streamlit as st
import pandas as pd
import io

# Webseiten-Konfiguration
st.set_page_config(page_title="Gnetz Mobilfunk Tracker", layout="wide")
st.title("🌐 Gnetz Team-Projekt-Tracker")

# 2. PFAD-ERKENNUNG (Für dich lokal am PC)
USER_PROFILE = os.environ.get("USERPROFILE", "")
EXCEL_PATH = os.path.join(USER_PROFILE, "OneDrive", "Gnetz neu", "1PROJEKTLISTE_GNETZneu.xlsx")

# Sicherheits-Suche für dich lokal am PC
if not os.path.exists(EXCEL_PATH) and USER_PROFILE:
    try:
        for folder in os.listdir(USER_PROFILE):
            if "OneDrive" in folder or "Ümit" in folder or "Persönlich" in folder:
                test_path = os.path.join(USER_PROFILE, folder, "Gnetz neu", "1PROJEKTLISTE_GNETZneu.xlsx")
                if os.path.exists(test_path):
                    EXCEL_PATH = test_path
                    break
    except Exception:
        pass

# 3. ENTSCHEIDUNG: Läuft die App bei dir lokal oder in der Cloud?
df = None
is_cloud = not os.path.exists(EXCEL_PATH)

if is_cloud:
    st.info("☁️ **Cloud-Modus für Kollegen aktiv.**")
    uploaded_file = st.file_uploader("Bitte ziehe die Datei '1PROJEKTLISTE_GNETZneu.xlsx' hier hinein:", type=["xlsx"])
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Fehler beim Laden: {e}")
else:
    st.success("💻 **Lokaler PC-Modus aktiv (Live-OneDrive-Anbindung).**")
    try:
        df = pd.read_excel(EXCEL_PATH)
    except Exception:
        st.error("🔒 Bitte schließe die originale Excel-Datei auf deinem PC!")

# 4. TABELLE UND AUTOMATISIERUNG
if df is not None:
    if "Abgehakt" not in df.columns:
        df["Abgehakt"] = False
    df = df.fillna("")

    # Interaktiver Tabellen-Editor
    edited_df = st.data_editor(
        df,
        hide_index=True,
        column_config={"Abgehakt": st.column_config.CheckboxColumn("Abgehakt", default=False)},
        use_container_width=True
    )

    # 5. SPEICHERN / EXPORTIEREN
    if not is_cloud:
        # Lokaler PC-Modus speichert direkt zurück in dein OneDrive
        if st.button("💾 Änderungen direkt in Excel speichern", type="primary"):
            try:
                edited_df.to_excel(EXCEL_PATH, index=False)
                st.success("🎉 Änderungen direkt in OneDrive-Excel gespeichert!")
            except PermissionError:
                st.error("🔒 Bitte schließe Excel auf deinem PC!")
    else:
        # Cloud-Modus für Kollegen generiert eine fertige Download-Datei
        st.write("👉 Klicke hier, um die aktualisierte Liste herunterzuladen und in OneDrive zu ersetzen:")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            edited_df.to_excel(writer, index=False)
        output.seek(0)
        st.download_button(
            label="📥 Aktualisierte Excel-Datei herunterladen",
            data=output,
            file_name="1PROJEKTLISTE_GNETZneu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    # Fortschrittsanzeige
    st.divider()
    erledigt = (edited_df["Abgehakt"] == True).sum()
    gesamt = len(edited_df)
    prozent = int((erledigt / gesamt) * 100) if gesamt > 0 else 0
    st.progress(prozent)
    st.write(f"**{erledigt} von {gesamt}** Projekten erledigt ({prozent}%).")
