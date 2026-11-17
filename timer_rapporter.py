"""Bygger og sender ugentlig Excel-rapport med ikke godkendte timer pr. leder."""
import os
import smtplib
from datetime import datetime
from email.message import EmailMessage

import pandas as pd
import win32com.client as win32

SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
SMTP_PORT = 25
AFSENDER = "RPA_info@aarhus.dk"
MODTAGER_KONSTANT = "MTMTimerRapportModtagere"

# Mulige navne på kolonnen med medarbejdernavn i SAP-udtrækket (første match bruges).
# Står det rigtige navn ikke her, så tilføj det øverst.
NAVN_KOLONNER = [
    "medarbejder_navn",  # KaldeNavn fra [ORG].[adm].[Bruger_Personaleleder]
    "Medarbejdernavn",
    "Navn",
    "Navn på medarbejder",
    "Medarbejder/ansøgernavn",
    "Medarb.-/ansøgernavn",
    "Formateret navn på medarbejder eller ansøger",
]

TAL_FORMAT = "#,##0.0"  # én decimal

# Excel-konstanter
XL_DATABASE = 1
XL_ROW_FIELD = 1
XL_SUM = -4157
XL_DESCENDING = 2
XL_DECIMAL_SEPARATOR = 3
XL_THOUSANDS_SEPARATOR = 4


def _find_navnekolonne(kolonner) -> str | None:
    normaliseret = {str(k).strip().lower(): k for k in kolonner}
    for kandidat in NAVN_KOLONNER:
        if kandidat.lower() in normaliseret:
            return normaliseret[kandidat.lower()]
    return None


def lav_pivot_rapport(detaljer: pd.DataFrame, sti: str) -> str | None:
    """Skriver data og bygger en pivottabel: Leder -> Medarbejder, sorteret efter flest timer."""
    # Fjern totalrækken fra SAP-udtrækket (den har intet medarbejdernummer)
    detaljer = detaljer[detaljer["medarbejder_id"].notna()]

    navn_kol = _find_navnekolonne(detaljer.columns)
    if navn_kol is None:
        print(f"⚠  Fandt ingen navnekolonne. Kolonner i udtrækket: {list(detaljer.columns)}")
        navn = pd.Series("(ukendt navn)", index=detaljer.index)
    else:
        print(f"Bruger '{navn_kol}' som medarbejdernavn")
        navn = detaljer[navn_kol].fillna("(ukendt navn)").astype(str).str.strip()

    az = detaljer["az"].fillna("ingen az").astype(str)

    data = pd.DataFrame({
        "Leder": detaljer["leder_kaldenavn"].fillna("(Ingen leder fundet)"),
        "Medarbejder": navn + " (" + az + ")",
        # Rund de enkelte rækker til 2 decimaler (fjerner støj som 7.3999999),
        # så kvarterer (0,25) stadig summeres præcist. Visningen er 1 decimal.
        "Timer": detaljer["ikke_reg_timer"].fillna(0).round(2),
    })
    data = data[data["Timer"] != 0]
    if data.empty:
        print("Ingen ikke godkendte timer – ingen rapport")
        return None

    data.to_excel(sti, sheet_name="Data", index=False)
    kilde = f"Data!R1C1:R{len(data) + 1}C{len(data.columns)}"

    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        # Pivotfeltets talformat læses med Excels lokale separatorer (dansk: "." og ","),
        # så formatet bygges ud fra maskinens egne indstillinger.
        intl = xl.International  # pywin32 returnerer alle indstillinger som en tuple (0-indekseret)
        if callable(intl):
            dec, tus = intl(XL_DECIMAL_SEPARATOR), intl(XL_THOUSANDS_SEPARATOR)
        else:
            dec, tus = intl[XL_DECIMAL_SEPARATOR - 1], intl[XL_THOUSANDS_SEPARATOR - 1]
        lokalt_format = f"#{tus}##0{dec}0"

        wb = xl.Workbooks.Open(os.path.abspath(sti))
        ws_data = wb.Worksheets("Data")
        ws_data.Columns("C").NumberFormat = TAL_FORMAT
        ws_pivot = wb.Worksheets.Add(Before=ws_data)
        ws_pivot.Name = "Timer pr. leder"

        cache = wb.PivotCaches().Create(SourceType=XL_DATABASE, SourceData=kilde)
        pt = cache.CreatePivotTable(TableDestination=ws_pivot.Range("A3"), TableName="TimerPrLeder")

        leder = pt.PivotFields("Leder")
        leder.Orientation = XL_ROW_FIELD
        leder.Position = 1
        medarb = pt.PivotFields("Medarbejder")
        medarb.Orientation = XL_ROW_FIELD
        medarb.Position = 2

        sum_felt = pt.AddDataField(pt.PivotFields("Timer"), "Ikke godkendte timer", XL_SUM)

        leder.AutoSort(XL_DESCENDING, "Ikke godkendte timer")
        medarb.AutoSort(XL_DESCENDING, "Ikke godkendte timer")
        # Sæt formatet igen efter layout/sortering, både på feltet og på de synlige celler
        sum_felt.NumberFormat = lokalt_format
        pt.DataBodyRange.NumberFormatLocal = lokalt_format
        leder.ShowDetail = False  # alle ledere foldet sammen fra start

        ws_pivot.Range("A1").Value = f"Ikke godkendte timer pr. {datetime.today():%d-%m-%Y}"
        ws_pivot.Activate()
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()
    return sti


def send_timer_rapport(orchestrator_connection, sti: str) -> None:
    raw = orchestrator_connection.get_constant(MODTAGER_KONSTANT).value
    modtagere = [m.strip() for m in raw.replace(";", ",").split(",") if m.strip()]
    if not modtagere:
        raise ValueError(f"Konstanten {MODTAGER_KONSTANT} indeholder ingen adresser")

    uge = datetime.today().isocalendar().week
    msg = EmailMessage()
    msg["From"] = AFSENDER
    msg["To"] = ", ".join(modtagere)
    msg["Subject"] = f"Ikke godkendte timer pr. leder – uge {uge}"
    msg.set_content("Please enable HTML to view this message.")
    msg.add_alternative(
        "<html><body><p>Hej</p>"
        "<p>Vedhæftet er oversigten over ikke godkendte timer pr. leder. "
        "Vh. robotten :)</p></body></html>",
        subtype="html",
    )
    with open(sti, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(sti),
        )
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as smtp:
        smtp.send_message(msg)
    print(f"✅ Timerrapport sendt til {len(modtagere)} modtagere")