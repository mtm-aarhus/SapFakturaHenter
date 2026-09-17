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
        {"RunName": 'SD løn udtræk', "UploadMappe": "SP"},
        {"RunName": "MTMIkkeGodkendteTimer", "UploadMappe": "SP"},
        {"RunName": "ZPSA_Brugerparametre", "UploadMappe": "SP"},
        {"RunName": "SD Forfaldne faktura", "UploadMappe": "SP"},
        {"RunName": "SD Stamdatatabel", "UploadMappe": "SP"},
        {"RunName": "SDAfstemning", "UploadMappe": "SP"},
        {"RunName": "KEX5", "UploadMappe": "SP"},
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
                sap_running = initialize_sap(orchestrator_connection)
                if not sap_running:
                    raise Exception("SAP failed to launch successfully")
                else:
                    print("SAP is running and ready.")
                watcher = start_popup_watcher(interval=0.3)
                try:
                    print("▶ Starter MTM ikke godkendte timer")
                    MTMIkkeGodkendteTimer()
                finally:
                    watcher.stop()
        
                cwd = os.getcwd()
                os.rename("ikkegodkendtetimer.XLSX", "MTMIkkeGodkendteTimer.xlsx")
                filepath = os.path.join(cwd, "MTMIkkeGodkendteTimer.xlsx")
        
                # # Nærmeste leder (Opus + ORG ligger på samme server -> én forbindelse)
                # sql_server_f = orchestrator_connection.get_constant("sqlserverf").value
                # conn_string_f = f"DRIVER={{SQL Server}};SERVER={sql_server_f};DATABASE=FDW;Trusted_Connection=yes;"
                # conn_f = pyodbc.connect(conn_string_f)
                # try:
                #     resultat = timerPerLeder(conn_f)
                # finally:
                #     conn_f.close()
        
                # # resultat er nu en DataFrame med leder + summerede timer
                # print(resultat)
                upload_to_sharepoint(Client, filepath, parent_folder_url, site_url_str=sharepoint_site_url)
                file_deleter(filepath)
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
                filepath_xlsx = os.path.join(cwd, "Opusbrugere.xlsx")
                filepath_html = os.path.join(cwd, "Opusbrugere.html")
                upload_to_sharepoint(Client, filepath_xlsx, parent_folder_url, site_url_str=sharepoint_site_url)
                file_deleter(filepath_xlsx)
                file_deleter(filepath_html)

                # file_deleter(filepath)
            except Exception as e:
                close_all_sap()
                orchestrator_connection.log_error(f'ZPSA_Brugerparametre fejlede {e}')
                continue

def timerPerLeder(conn_org):
    # 1. Excel-data
    df = pd.read_excel("MTMIkkeGodkendteTimer.xlsx")
    df = df.rename(columns={
        "Medarbejdernummer": "medarbejder_id",
        "Antal (måleenhed)": "ikke_reg_timer",
    })

    # Konverter timer til tal (fanger tekst-tal som '5,0' eller ' 5 ')
    df["ikke_reg_timer_raw"] = df["ikke_reg_timer"]
    df["ikke_reg_timer"] = pd.to_numeric(df["ikke_reg_timer"], errors="coerce")
    total_excel = df["ikke_reg_timer"].sum()

    # 2. Oversæt medarbejdernummer -> az (Ident) i Opus-db
    medarbejdere = df["medarbejder_id"].dropna().unique().tolist()
    placeholders = ",".join("?" * len(medarbejdere))
    query_az = f"""
        SELECT MedarbejderNummer AS medarbejder_id,
               Ident             AS az
        FROM [Opus].[brugerstyring].[BRS_Rolletildeling-Hist]
        WHERE MedarbejderNummer IN ({placeholders})
    """
    az_map = pd.read_sql(query_az, conn_org, params=medarbejdere)

    # Behold kun én az pr. medarbejdernummer
    az_map = az_map.drop_duplicates(subset="medarbejder_id", keep="first")

    df = df.merge(az_map, on="medarbejder_id", how="left")

    # 3. Hent leder-info for alle az'er i ORG-db
    azer = df["az"].dropna().unique().tolist()
    placeholders = ",".join("?" * len(azer))
    query_leder = f"""
        SELECT BrugerNavn                AS az,
               FungerendeLederBrugernavn AS leder_brugernavn,
               FungerendeLederKaldenavn  AS leder_kaldenavn,
               FungerendeLederEmail      AS leder_email
        FROM [ORG].[adm].[Bruger_AD_PrimærKonto_Aktuel]
        WHERE BrugerNavn IN ({placeholders})
    """
    ledere = pd.read_sql(query_leder, conn_org, params=azer)

    # Behold kun én lederrække pr. az
    ledere = ledere.drop_duplicates(subset="az", keep="first")
    df = df.merge(ledere, on="az", how="left")

        # 4. Sum ikke-reg. timer pr. leder
    resultat = (
        df.groupby(
            ["leder_brugernavn", "leder_kaldenavn", "leder_email"],
            dropna=False,
        )["ikke_reg_timer"]
        .sum()
        .reset_index()
        .sort_values("ikke_reg_timer", ascending=False)
    )

    diff = resultat["ikke_reg_timer"].sum() - total_excel
    if abs(diff) > 0.01:
        print(f"⚠  DIFFERENCE: {diff}  (resultat matcher IKKE Excel-total!)")
    else:
        print("✅ Total matcher Excel-total")

    # Skriv resultat til CSV (dansk Excel: semikolon-separator, komma-decimal)
    csv_path = os.path.join(os.getcwd(), "TimerPerLeder.csv")
    resultat.to_csv(
        csv_path,
        sep=";",
        decimal=",",
        index=False,
        encoding="utf-8-sig",   # -sig sikrer at æ/ø/å vises rigtigt i Excel
    )

    return resultat
