import os, time, shutil, tempfile
import random
import string
from datetime import datetime, timedelta
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from urllib.parse import urlparse
from email.message import EmailMessage
import smtplib
import mimetypes
from robot_framework.initialize_sap import initialize_sap
from scripts import *
from sap_popup_utils import start_popup_watcher
import os, time, shutil, tempfile, mimetypes
from email.message import EmailMessage
import smtplib
from openpyxl import load_workbook, Workbook
import pyodbc
import pandas as pd

# Opsæt connection til Orchestrator
orchestrator_connection = OrchestratorConnection(
    "SapProcess",
    os.getenv('OpenOrchestratorSQL'),
    os.getenv('OpenOrchestratorKey'),
    None)

def process(orchestrator_connection: OrchestratorConnection) -> None:
    
    def Email(Modtagermail, Bcc1, Bcc2, file_name, file_path):
        SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
        SMTP_PORT = 25
        subject = "Indtastningsgrundlag i forhold til SD Løn"

        html = """
        <html>
        <body>
            <p>Hej HR:) </p>
            <p>Hermed som aftalt indtastningsgrundlag i forhold til SD Løn.</p> 
        </body>
        </html>
        """

        # Sørg for streng-path og midlertidig kopi for at undgå låse
        src = str(file_path)
        base = os.path.basename(src)
        tmp_dir = tempfile.gettempdir()
        tmp_path = os.path.join(tmp_dir, f"mail_{int(time.time()*1000)}_{base}")

        # kopi med lille backoff hvis OneDrive/AV holder et håndtag
        delay = 0.3
        for attempt in range(1, 6):
            try:
                shutil.copyfile(src, tmp_path)
                break
            except PermissionError:
                if attempt == 5:
                    raise
                time.sleep(delay)
                delay *= 1.7

        msg = EmailMessage()
        msg["To"] = Modtagermail
        msg["From"] = 'RPA_info@aarhus.dk'
        msg["Subject"] = subject
        msg["Cc"] = orchestrator_connection.get_constant('Error Email').value
        msg["Bcc"] = ", ".join(filter(None, [Bcc1, Bcc2]))
        msg.set_content("Please enable HTML to view this message.")
        msg.add_alternative(html, subtype="html")

        try:
            # Vedhæft fra midlertidig, ikke fra OneDrive-sti
            mime_type, _ = mimetypes.guess_type(tmp_path)
            maintype, subtype = mime_type.split("/") if mime_type else ("application", "octet-stream")
            with open(tmp_path, "rb") as f:
                msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=file_name or base)
        except Exception as e:
            print(f"Fejl under vedhæftning af fil: {e}")
            try:
                os.remove(tmp_path)
            except OSError:
                pass
            raise

        # Send
        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as smtp:
                smtp.send_message(msg)
                print("✅ Mail sendt")
        except Exception as e:
            print(f"❌ Failed to send email: {e}")
            raise
        finally:
            try:
                os.remove(tmp_path)
            except OSError:
                pass

    def file_deleter(filename):
        if os.path.exists(filename):
            os.remove(filename)
        else:
            print("The file does not exist")

    def sharepoint_client(site_url) -> ClientContext:

        certification = orchestrator_connection.get_credential("SharePointCert")
        api = orchestrator_connection.get_credential("SharePointAPI")

        cert_credentials = {
            "tenant": api.username,
            "client_id": api.password,
            "thumbprint": certification.username,
            "cert_path": certification.password
        }

        ctx = ClientContext(site_url).with_client_certificate(**cert_credentials)

        return ctx

    def _server_relative(folder_url: str, site_url_str: str) -> str:
        """
        Returnér en server-relativ sti (starter med /...).
        - folder_url kan være fuld https-URL, server-relativ (/teams/...), eller relativ ('Delte dokumenter/X').
        - site_url_str er den STRÆNG du gav til ClientContext(...).
        """
        if not isinstance(site_url_str, str):
            site_url_str = str(site_url_str)

        # Fuld URL -> strip domæne, behold path
        if isinstance(folder_url, str) and folder_url.lower().startswith("http"):
            path = urlparse(folder_url).path
        else:
            path = folder_url

        if not isinstance(path, str):
            path = str(path)

        if path.startswith("/"):
            return path  # allerede server-relativ

        # Relativ sti -> præfikser med web-rodens path
        base_path = urlparse(site_url_str).path.rstrip("/")
        if base_path:
            return f"{base_path}/{path}".replace("\\", "/")
        else:
            return f"/{path}".replace("\\", "/")

    def upload_to_sharepoint(ctx, file_path, folder_url: str, site_url_str: str, max_retries: int = 6):

        file_path = str(file_path)
        file_name = os.path.basename(file_path)

        # Midlertidig kopi at uploade fra
        tmp_dir = tempfile.gettempdir()
        tmp_path = os.path.join(tmp_dir, f"upload_{int(time.time()*1000)}_{file_name}")

        # Kopi med backoff (hvis OneDrive/AV holder håndtag i millisekunder)
        delay = 0.4
        for attempt in range(1, max_retries + 1):
            try:
                shutil.copyfile(file_path, tmp_path)
                break
            except PermissionError as e:
                if attempt == max_retries:
                    raise
                time.sleep(delay)
                delay *= 1.7

        # Normaliser målmappe som server-relativ sti
        srv_rel = _server_relative(folder_url, site_url_str)
        target_folder = ctx.web.get_folder_by_server_relative_url(srv_rel)

        # Upload med retries
        delay = 0.4
        for attempt in range(1, max_retries + 1):
            try:
                with open(tmp_path, "rb") as f:
                    content = f.read()
                target_folder.upload_file(file_name, content)
                ctx.execute_query()
                print(f"✅ Uploaded: {file_name} -> {srv_rel}")
                break
            except PermissionError as e:
                if attempt == max_retries:
                    raise
                time.sleep(delay)
                delay *= 1.7

        try:
            os.remove(tmp_path)
        except OSError:
            pass

    sharepoint_site_url = orchestrator_connection.get_constant('AarhusKommuneSharePoint').value
    sharepoint_site_url = f'{sharepoint_site_url}/Teams/tea-teamsite10343'
    parent_folder_url = sharepoint_site_url.split(".com")[-1] +'/Delte Dokumenter/Dataprojekt/2026'
    Client = sharepoint_client( site_url= sharepoint_site_url)

    runs = [
        # {"RunName": 'SD løn udtræk', "UploadMappe": "SP"},
        {"RunName": "MTMIkkeGodkendteTimer", "UploadMappe": "SP"},
        # {"RunName": "ZPSA_Brugerparametre", "UploadMappe": "SP"}
        # {"RunName": "SD Forfaldne faktura", "UploadMappe": "SP"},
        # {"RunName": "SD Stamdatatabel", "UploadMappe": "SP"},
        # {"RunName": "SDAfstemning", "UploadMappe": "SP"},
        # {"RunName": "KEX5", "UploadMappe": "SP"},
    ]

    for run in runs:
        if run["RunName"] == "SD løn udtræk" and datetime.today().weekday() == 0:
            try:
                sap_running = initialize_sap(orchestrator_connection)
                if not sap_running:
                    raise Exception("SAP failed to launch successfully")
                else:
                    print("SAP is running and ready.")
                watcher = start_popup_watcher(interval= 0.3)
                try:
                    print("▶ Starter SD løn udtræk")
                    SDLonUdtrak()
                    
                finally:
                    watcher.stop()
                Outfile, Name = InputToTemplate()

                # upload_to_sharepoint(Client, Outfile, parent_folder_url, site_url_str=sharepoint_site_url) ##skal ikke aktiveres for nu
                Mail = orchestrator_connection.get_constant('SapFakturaHenterHRMail').value
                ModtagerMail = Mail.split(',')[0]
                Bcc1 = Mail.split(',')[-1]
                Bcc2 = Mail.split(',')[1]
                Email(ModtagerMail, Bcc1, Bcc2, file_name= Name, file_path= Outfile)
                file_deleter(Outfile)
                file_deleter('export.xlsx')
            except Exception as e:
                close_all_sap()
                orchestrator_connection.log_error(f'SD løn udtræk fejlede {e} - genkør!!')
                continue

        elif run["RunName"] == "KEX5":
            cwd = os.getcwd()
            combined_path = os.path.join(cwd, "KE5x_samlet.xlsx")

            ke5x_runs = [
                {"title": "title2.xlsx",  "dst": os.path.join(cwd, "KE5x_2.xlsx")},
                {"title": "title82.xlsx", "dst": os.path.join(cwd, "KE5x_82.xlsx")},
            ]

            def _resolve_ke5x_output(base, title):
                # KE5x gemmer uden fast endelse; prøv de mest sandsynlige varianter
                for cand in (f"{title}.XLSX", f"{title}.xlsx", title):
                    p = os.path.join(base, cand)
                    if os.path.exists(p):
                        return p
                raise FileNotFoundError(
                    f"Kunne ikke finde KE5x-output for '{title}' i {base} "
                    f"(prøvede {title}.XLSX / .xlsx / uden endelse)"
                )

            saved_files = []
            try:
                for r in ke5x_runs:
                    sap_running = initialize_sap(orchestrator_connection)
                    if not sap_running:
                        raise Exception("SAP failed to launch successfully")
                    print("SAP is running and ready.")

                    watcher = start_popup_watcher(interval=0.3)
                    try:
                        print(f"▶ Starter KE5x ({r['title']})")
                        KE5x(orchestrator_connection, r["title"])
                    finally:
                        watcher.stop()

                    # KE5x kalder selv close_all_sap(), så SAP er lukket her.
                    # Flyt output væk med det samme under et entydigt navn.
                    produced = _resolve_ke5x_output(cwd, r["title"])
                    if os.path.exists(r["dst"]):
                        os.remove(r["dst"])
                    os.rename(produced, r["dst"])
                    saved_files.append(r["dst"])

                # Stabl de to rapporter til én tabel (behold kun header fra første fil)
                combined_wb = Workbook()
                combined_ws = combined_wb.active
                combined_ws.title = "KE5x"
                for i, f in enumerate(saved_files):
                    wb = load_workbook(f, read_only=True, data_only=True)
                    ws = wb.active
                    min_row = 1 if i == 0 else 2
                    for row in ws.iter_rows(min_row=min_row, values_only=True):
                        combined_ws.append(row)
                    wb.close()

                combined_wb.save(combined_path)
                print(f"✅ Samlet {len(saved_files)} rapporter til "
                    f"{os.path.basename(combined_path)} ({combined_ws.max_row} rækker inkl. header)")

                upload_to_sharepoint(Client, combined_path, parent_folder_url, site_url_str=sharepoint_site_url)
                for f in saved_files:
                    file_deleter(f)
                file_deleter(combined_path)

            except Exception as e:
                close_all_sap()
                for f in saved_files + [combined_path]:
                    try:
                        if os.path.exists(f):
                            os.remove(f)
                    except OSError:
                        pass
                orchestrator_connection.log_error(f'KEX5 {e} ')
                continue

        elif run["RunName"] == "SD Forfaldne faktura":
            try:
                sap_running = initialize_sap(orchestrator_connection)
                if not sap_running:
                    raise Exception("SAP failed to launch successfully")
                else:
                    print("SAP is running and ready.")
                watcher = start_popup_watcher(interval= 0.3)
                try:
                    print("▶ Starter SD løn udtræk")
                    SDForfaldneFaktura(orchestrator_connection)
                    
                finally:
                    watcher.stop()

                cwd = os.getcwd()
                filepath = os.path.join(cwd, "Forfaldne fakturaer MTM.XLSX")

                upload_to_sharepoint(Client, filepath, parent_folder_url, site_url_str=sharepoint_site_url)
                file_deleter(filepath)
            except Exception as e:
                close_all_sap()
                orchestrator_connection.log_error(f'SD forfaldne faktura fejlede {e} ')
                continue

        elif run["RunName"] == "SD Stamdatatabel":
            try:
                sap_running = initialize_sap(orchestrator_connection)
                if not sap_running:
                    raise Exception("SAP failed to launch successfully")
                else:
                    print("SAP is running and ready.")
                watcher = start_popup_watcher(interval= 0.3)
                try:
                    print("▶ Starter SD løn udtræk")
                    SDStamdataTabel(orchestrator_connection)
                    
                finally:
                    watcher.stop()

                cwd = os.getcwd()
                filepath = os.path.join(cwd, "Stamdatatabel.XLSX")

                upload_to_sharepoint(Client, filepath, parent_folder_url, site_url_str=sharepoint_site_url)
                file_deleter(filepath)
            except Exception as e:
                close_all_sap()
                orchestrator_connection.log_error(f'SD stamdata fejlede {e} ')
                continue
        elif run["RunName"] == "SDAfstemning":
            try:
                sap_running = initialize_sap(orchestrator_connection)
                if not sap_running:
                    raise Exception("SAP failed to launch successfully")
                else:
                    print("SAP is running and ready.")
                watcher = start_popup_watcher(interval= 0.3)
                try:
                    print("▶ Starter SD afstemning")
                    SDAfstemning()

                finally:
                    watcher.stop()

                cwd = os.getcwd()
                filepath = os.path.join(cwd, "Opus data til afst.xlsx")
                upload_to_sharepoint(Client, filepath, parent_folder_url, site_url_str=sharepoint_site_url)
                file_deleter(filepath)
            except Exception as e:
                close_all_sap()
                orchestrator_connection.log_error(f'SD afstemning fejlede {e}')
                continue

        elif run["RunName"] == "MTMIkkeGodkendteTimer":
            try:
                # sap_running = initialize_sap(orchestrator_connection)
                # if not sap_running:
                #     raise Exception("SAP failed to launch successfully")
                # else:
                #     print("SAP is running and ready.")
                # watcher = start_popup_watcher(interval=0.3)
                # try:
                #     print("▶ Starter MTM ikke godkendte timer")
                #     MTMIkkeGodkendteTimer()
                # finally:
                #     watcher.stop()

                cwd = os.getcwd()
                # os.rename("ikkegodkendtetimer.XLSX", "MTMIkkeGodkendteTimer.xlsx")
                filepath = os.path.join(cwd, "MTMIkkeGodkendteTimer.xlsx")

                # Nærmeste leder (Opus + ORG ligger på samme server -> én forbindelse)
                sql_server_f = orchestrator_connection.get_constant("sqlserverf").value
                conn_string_f = f"DRIVER={{SQL Server}};SERVER={sql_server_f};DATABASE=FDW;Trusted_Connection=yes;"
                conn_f = pyodbc.connect(conn_string_f)
                try:
                    resultat = timerPerLeder(conn_f)
                finally:
                    conn_f.close()

                # resultat er nu en DataFrame med leder + summerede timer
                print(resultat)
                # file_deleter(filepath)
            except Exception as e:
                close_all_sap()
                orchestrator_connection.log_error(f'MTM ikke godkendte timer fejlede {e}')
                continue

        elif run["RunName"] == "ZPSA_Brugerparametre":
            try:
                sap_running = initialize_sap(orchestrator_connection)
                if not sap_running:
                    raise Exception("SAP failed to launch successfully")
                else:
                    print("SAP is running and ready.")
                watcher = start_popup_watcher(interval=0.3)
                try:
                    print("▶ Starter ZPSA_Brugerparametre")
                    ZPSA_Brugerparametre()
                finally:
                    watcher.stop()
            
                cwd = os.getcwd()
                filepath = os.path.join(cwd, "Opusbrugere.txt")

                # file_deleter(filepath)
            except Exception as e:
                close_all_sap()
                orchestrator_connection.log_error(f'ZPSA_Brugerparametre fejlede {e}')
                continue

import unicodedata

def timerPerLeder(conn_org):
    print("=" * 60)
    print("TIMER PER LEDER - DIAGNOSTIK")
    print("=" * 60)

    # --- DEBUG: sæt medarbejdernummer + az ind for at spore én medarbejder ---
    NR = None        # fx 3318  (medarbejdernummer, int)
    AZ = None        # fx "aztstn2"  (az i lowercase - matcher efter casefold)

    # Normaliser tekst så ø/encoding, whitespace og case ikke bryder sammenligning
    def _norm(s):
        if pd.isna(s):
            return ""
        return unicodedata.normalize("NFC", str(s)).strip().casefold()

    # 1. Excel-data
    df = pd.read_excel("MTMIkkeGodkendteTimer.xlsx")
    df = df.rename(columns={
        "Medarbejdernummer": "medarbejder_id",
        "Antal (måleenhed)": "ikke_reg_timer",
    })

    # Fjern rækker UDEN medarbejder_id (sum-/subtotalrække nederst i arket)
    før = len(df)
    df = df[df["medarbejder_id"].notna()].copy()
    df = df[df["medarbejder_id"].astype(str).str.strip() != ""].copy()
    if før - len(df):
        print(f"⚠  Fjernede {før - len(df)} række(r) uden medarbejder_id (sum-/tomme rækker)")

    # Excel læser medarbejder_id som float (3853.0) pga. tomme celler -> match fejler.
    # Normaliser til heltal (nullable Int64) i BEGGE sider af merget.
    df["medarbejder_id"] = pd.to_numeric(df["medarbejder_id"], errors="coerce").astype("Int64")
    df["ikke_reg_timer"] = pd.to_numeric(df["ikke_reg_timer"], errors="coerce")
    total_excel = df["ikke_reg_timer"].sum()

    print("\n--- 1. EXCEL ---")
    print(f"Rækker efter filtrering:   {len(df)}")
    print(f"Unikke medarbejdere:       {df['medarbejder_id'].nunique()}")
    print(f"TIMER I ALT (rå Excel):    {total_excel}")

    # 2. Oversæt medarbejdernummer -> az (Ident) i Opus-db  (BEHOLD ALLE az)
    medarbejdere = [int(x) for x in df["medarbejder_id"].dropna().unique()]
    placeholders = ",".join("?" * len(medarbejdere))
    query_az = f"""
        SELECT MedarbejderNummer AS medarbejder_id,
               Ident             AS az
        FROM [Opus].[brugerstyring].[BRS_Rolletildeling]
        WHERE MedarbejderNummer IN ({placeholders})
    """
    az_map = pd.read_sql(query_az, conn_org, params=medarbejdere)
    az_map = az_map.drop_duplicates()   # fjern kun identiske gentagelser
    az_map["medarbejder_id"] = pd.to_numeric(az_map["medarbejder_id"], errors="coerce").astype("Int64")
    # Casefold az: Opus giver lowercase, ORG-viewet uppercase -> ellers matcher join ikke
    az_map["az"] = az_map["az"].astype(str).str.strip().str.casefold()

    print("\n--- 2. AZ-OPSLAG (Opus) ---")
    print(f"Rækker efter drop_duplicates: {len(az_map)}")
    print(f"Unikke medarbejdere:          {az_map['medarbejder_id'].nunique()}")
    print(f"Unikke az:                    {az_map['az'].nunique()}")

    # DEBUG A: er NR + az kommet med i az_map?
    if NR is not None:
        print(f"A. NR i az_map:   {(az_map['medarbejder_id'] == NR).any()}")
        print(f"   az i az_map:   {az_map.loc[az_map['medarbejder_id'] == NR, 'az'].tolist()}")

    # 3. Hent leder-info + Niveau2 UDEN at filtrere i SQL
    #    (TM-filtrering sker i pandas -> undgår ø/encoding-mismatch mod databasen)
    azer = az_map["az"].dropna().unique().tolist()
    # BrugerNavn i basen er uppercase; send uppercase-varianter som parametre (robust mod case-sensitiv collation)
    azer_db = [a.upper() for a in azer]
    placeholders = ",".join("?" * len(azer_db))
    query_leder = f"""
        SELECT BrugerNavn                AS az,
               Niveau2_OrgNavn           AS niveau2,
               FungerendeLederBrugernavn AS leder_brugernavn,
               FungerendeLederKaldenavn  AS leder_kaldenavn,
               FungerendeLederEmail      AS leder_email
        FROM [ORG].[adm].[Bruger_AD_PrimærKonto_Aktuel]
        WHERE BrugerNavn IN ({placeholders})
    """
    ledere_alle = pd.read_sql(query_leder, conn_org, params=azer_db)

    # Casefold az på ORG-siden så join matcher az_map
    ledere_alle["az"] = ledere_alle["az"].astype(str).str.strip().str.casefold()

    # Filtrer til Teknik og Miljø på normaliseret tekst
    mask_tm = ledere_alle["niveau2"].map(_norm) == _norm("Teknik og Miljø")
    ledere_tm = ledere_alle[mask_tm].drop(columns="niveau2").drop_duplicates()
    tm_az = set(ledere_tm["az"])

    # DEBUG B: overlever az'en TM-filteret?
    if AZ is not None:
        print(f"B. AZ i ledere_tm (TM-filtreret): {(ledere_tm['az'] == AZ).any()}")

    # ALLE az'er for medarbejdere med mere end 1 az - med TM-markering
    az_pr_medarb = az_map.groupby("medarbejder_id")["az"].nunique()
    flere_az = az_pr_medarb[az_pr_medarb > 1]
    if len(flere_az):
        vis = az_map[az_map["medarbejder_id"].isin(flere_az.index)].copy()
        vis["i_TM"] = vis["az"].isin(tm_az)
        print(f"\n⚠  {len(flere_az)} medarbejder(e) med flere az - ALLE fundne az'er:")
        print(vis.sort_values(["medarbejder_id", "az"]).to_string(index=False))
    else:
        print("\nIngen medarbejdere med flere az.")

    print("\n--- 3. LEDER-OPSLAG (ORG, kun Teknik og Miljø) ---")
    print(f"az før TM-filter:          {len(azer)}")
    print(f"az efter TM-filter:        {ledere_tm['az'].nunique()}")

    # Print de az'er der frasorteres af TM-filteret (ikke i Teknik og Miljø)
    fjernede = az_map[~az_map["az"].isin(tm_az)]
    if len(fjernede):
        fjernede = fjernede.merge(
            ledere_alle[["az", "niveau2"]].drop_duplicates(), on="az", how="left"
        )
        print(f"⚠  {fjernede['az'].nunique()} az frasorteret (ej Teknik og Miljø) - VÆRDIER:")
        print(fjernede.drop_duplicates(subset=["medarbejder_id", "az"])
              .sort_values("medarbejder_id").to_string(index=False))
    else:
        print("Ingen az frasorteret af TM-filteret.")

    # Kobl TM-az'er tilbage til medarbejder (inner => ikke-TM az'er falder fra)
    tm = az_map.merge(ledere_tm, on="az", how="inner")

    # DEBUG C: kom NR med igennem az_map->ledere_tm koblingen?
    if NR is not None:
        print(f"C. NR i tm (efter az->leder join): {(tm['medarbejder_id'] == NR).any()}")

    # 3b. TJEK: har et id STADIG mere end 1 az efter TM-filter? -> FEJL
    az_efter = tm.groupby("medarbejder_id")["az"].nunique()
    bad_ids = az_efter[az_efter > 1]
    if len(bad_ids):
        print(f"\n❌ {len(bad_ids)} medarbejder(e) har FLERE az selv efter TM-filter:")
        print(tm[tm["medarbejder_id"].isin(bad_ids.index)]
              [["medarbejder_id", "az", "leder_brugernavn"]]
              .sort_values("medarbejder_id").to_string(index=False))
        raise ValueError(
            "Flere az i Teknik og Miljø for medarbejder-id: "
            f"{bad_ids.index.tolist()} - kan ikke afgøre entydig az."
        )
    print("Alle medarbejdere har ≤1 az efter TM-filter ✅")

    # tm har nu ≤1 række pr. medarbejder_id -> merge giver INGEN fan-out
    df = df.merge(tm, on="medarbejder_id", how="left")

    # DEBUG D: hvad står der på NR efter det sidste merge?
    if NR is not None:
        print("D. NR efter df-merge:")
        print(df.loc[df["medarbejder_id"] == NR,
              ["medarbejder_id", "az", "leder_brugernavn"]].to_string(index=False))

    # Fjern rækker hvor personen er sin EGEN leder (az == leder_brugernavn)
    egen_leder = df["az"].notna() & df["leder_brugernavn"].notna() & (
        df["az"].astype(str).str.casefold()
        == df["leder_brugernavn"].astype(str).str.casefold()
    )
    if egen_leder.any():
        fjern = (df[egen_leder]
                 .groupby("az")["ikke_reg_timer"]
                 .agg(["size", "sum"])
                 .reset_index()
                 .rename(columns={"size": "antal_rækker", "sum": "timer"}))
        print(f"\n⚠  EGEN LEDER: fjerner {egen_leder.sum()} række(r) / "
              f"{df.loc[egen_leder, 'ikke_reg_timer'].sum()} timer "
              f"({fjern['az'].nunique()} person(er)):")
        print(fjern.sort_values("timer", ascending=False).to_string(index=False))
            # DEBUG: vis hvad der faktisk sammenlignes for de matchede
    if egen_leder.any():
        prøve = df.loc[egen_leder, ["medarbejder_id", "az", "leder_brugernavn"]].drop_duplicates()
        print("\nDEBUG egen-leder - faktiske værdier der matcher:")
        print(prøve.head(20).to_string(index=False))
        df = df[~egen_leder].copy()
    else:
        print("\nIngen personer er deres egen leder.")

    print(f"\nRækker efter merge:        {len(df)}")
    print(f"Rækker UDEN TM-leder:      {df['leder_brugernavn'].isna().sum()}")
    print(f"TIMER I ALT efter merge:   {df['ikke_reg_timer'].sum()}")

    # Medarbejdere hvis timer IKKE havner på en leder
    uden_leder = df[df["leder_brugernavn"].isna()]
    if len(uden_leder):
        opsummeret = (
            uden_leder.groupby("medarbejder_id")
            .agg(
                az=("az", "first"),
                antal_rækker=("ikke_reg_timer", "size"),
                timer=("ikke_reg_timer", "sum"),
            )
            .reset_index()
            .sort_values("timer", ascending=False)
        )
        print(f"\n⚠  {len(opsummeret)} medarbejder(e) UDEN leder - "
              f"{opsummeret['timer'].sum()} timer falder ud:")
        print(opsummeret.to_string(index=False))
    else:
        print("\nAlle medarbejdere koblet til en leder ✅")

    # 4. Sum ikke-reg. timer pr. leder (gruppér på brugernavn = stabil nøgle)
    resultat = (
        df.groupby("leder_brugernavn", dropna=False)
          .agg(
              leder_kaldenavn=("leder_kaldenavn", "first"),
              leder_email=("leder_email", "first"),
              ikke_reg_timer=("ikke_reg_timer", "sum"),
          )
          .reset_index()
          .sort_values("ikke_reg_timer", ascending=False)
    )

    print("\n--- 4. RESULTAT ---")
    print(f"Antal ledere i resultat:   {len(resultat)}")
    print(f"TIMER I ALT (resultat):    {resultat['ikke_reg_timer'].sum()}")
    print(f"TIMER I ALT (rå Excel):    {total_excel}")
    diff = resultat["ikke_reg_timer"].sum() - total_excel
    print(f"(Difference forventes pga. egen-leder + evt. umatchede: {diff})")

    print("\nTop 5 ledere:")
    print(resultat.head(5).to_string(index=False))

    # Tilføj total-række nederst (kun i CSV, ikke i returværdien)
    total_row = pd.DataFrame([{
        "leder_brugernavn": "TOTAL",
        "leder_kaldenavn": "",
        "leder_email": "",
        "ikke_reg_timer": resultat["ikke_reg_timer"].sum(),
    }])
    resultat_csv = pd.concat([resultat, total_row], ignore_index=True)

    # Skriv til CSV (dansk Excel: semikolon, komma-decimal)
    csv_path = os.path.join(os.getcwd(), "TimerPerLeder.csv")
    resultat_csv.to_csv(csv_path, sep=";", decimal=",", index=False, encoding="utf-8-sig")
    print(f"\n✅ Resultat skrevet til: {csv_path}")
    print("=" * 60)

    return resultat

process(orchestrator_connection= orchestrator_connection)



