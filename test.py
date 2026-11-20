"""
Test af timerrapporten (pivottabel: Leder -> Medarbejder -> Dato).

Kører:      SAP-udtræk (valgfrit) -> leder-opslag -> pivot-Excel -> mail til konstanten 'balas'
Kører IKKE: SharePoint-upload, køelementer, mail til de rigtige modtagere.
"""
import os
from datetime import datetime

import pyodbc
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

import timer_rapporter
from timer_rapporter import lav_pivot_rapport, send_timer_rapport
from robot_framework.process import timerPerLeder, hent_undtagne_ledere
from robot_framework.initialize_sap import initialize_sap
from scripts import MTMIkkeGodkendteTimer, close_all_sap
from sap_popup_utils import start_popup_watcher

# ---------------- Indstillinger ----------------
BRUG_EKSISTERENDE_FIL = True   # True: genbrug MTMIkkeGodkendteTimer.xlsx hvis den findes, ellers køres SAP
SEND_MAIL = False              # True: send rapporten til TEST_MODTAGER_KONSTANT
BEHOLD_RAPPORT = True          # True: lad rapporten ligge, så du kan åbne den bagefter
TEST_MODTAGER_KONSTANT = "balas"
# -----------------------------------------------

orchestrator_connection = OrchestratorConnection(
    "SapProcessTest",
    os.getenv("OpenOrchestratorSQL"),
    os.getenv("OpenOrchestratorKey"),
    None,
)

# Send kun til testkonstanten – den rigtige modtagerliste bliver aldrig læst
timer_rapporter.MODTAGER_KONSTANT = TEST_MODTAGER_KONSTANT

cwd = os.getcwd()
filepath = os.path.join(cwd, "MTMIkkeGodkendteTimer.xlsx")
rapport_sti = os.path.join(cwd, f"TEST_IkkeGodkendteTimer_{datetime.today():%Y-%m-%d}.xlsx")

# 1. Hent data fra SAP (eller genbrug fil)
if BRUG_EKSISTERENDE_FIL and os.path.exists(filepath):
    print(f"▶ Genbruger {filepath}")
else:
    try:
        if not initialize_sap(orchestrator_connection):
            raise Exception("SAP failed to launch successfully")
        watcher = start_popup_watcher(interval=0.3)
        try:
            print("▶ Starter MTM ikke godkendte timer")
            MTMIkkeGodkendteTimer()
        finally:
            watcher.stop()
        os.replace("ikkegodkendtetimer.XLSX", filepath)
    except Exception:
        close_all_sap()
        raise

# 2. Slå ledere op
sql_server_f = orchestrator_connection.get_constant("sqlserverf").value
conn_f = pyodbc.connect(
    f"DRIVER={{SQL Server}};SERVER={sql_server_f};DATABASE=FDW;Trusted_Connection=yes;"
)
try:
    resultat, detaljer = timerPerLeder(conn_f)
finally:
    conn_f.close()

print(f"\nKolonner i udtrækket: {list(detaljer.columns)}\n")

# 3. Vis hvad der VILLE være lagt i køen (intet lægges i kø)
undtagne = hent_undtagne_ledere(orchestrator_connection)
leder_az = resultat["leder_brugernavn"].fillna("").str.strip().str.lower()
til_koe = resultat[~leder_az.isin(undtagne)]
print(f"Undtagne ledere: {sorted(undtagne) or 'ingen'}")
print(f"Ledere i alt: {len(resultat)} – ville lægges i kø: {len(til_koe)} (gøres ikke i test)")
print(til_koe.head(15).to_string(index=False))
print()

# 4. Lav rapport og evt. send til testmodtager
try:
    if lav_pivot_rapport(detaljer, rapport_sti):
        print(f"✅ Rapport lavet: {rapport_sti}")
        if SEND_MAIL:
            send_timer_rapport(orchestrator_connection, rapport_sti)
        else:
            print("SEND_MAIL = False – ingen mail sendt")
finally:
    if not BEHOLD_RAPPORT and os.path.exists(rapport_sti):
        os.remove(rapport_sti)