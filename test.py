"""
Test af timerrapporten (pivottabel pr. leder).

Kører:   SAP-udtræk (valgfrit) -> leder-opslag -> pivot-Excel -> mail til konstanten 'balas'
Kører IKKE: SharePoint-upload, køelementer, mail til de rigtige modtagere.
"""
import os
from datetime import datetime

import pyodbc
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

import timer_rapporter
from timer_rapporter import lav_pivot_rapport, send_timer_rapport
from robot_framework.process import timerPerLeder
from robot_framework.initialize_sap import initialize_sap
from scripts import MTMIkkeGodkendteTimer, close_all_sap
from sap_popup_utils import start_popup_watcher

# ---------------- Indstillinger ----------------
BRUG_EKSISTERENDE_FIL = True   
SEND_MAIL = True               
BEHOLD_RAPPORT = False          
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

# 2. Slå ledere op (ingen kø)
sql_server_f = orchestrator_connection.get_constant("sqlserverf").value
conn_f = pyodbc.connect(
    f"DRIVER={{SQL Server}};SERVER={sql_server_f};DATABASE=FDW;Trusted_Connection=yes;"
)
try:
    resultat, detaljer = timerPerLeder(conn_f)
finally:
    conn_f.close()

print("\nDette VILLE være lagt i køen (gøres ikke i test):")
print(resultat.head(15).to_string(index=False))
print(f"... {len(resultat)} ledere i alt\n")
print(f"Kolonner i SAP-udtrækket: {list(detaljer.columns)}\n")

# 3. Lav rapport og send til testmodtager
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