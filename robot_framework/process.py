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
from scripts import SDLonUdtrak, InputToTemplate, SDForfaldneFaktura, SDStamdataTabel
from sap_popup_utils import start_popup_watcher
import os, time, shutil, tempfile, mimetypes
from email.message import EmailMessage
import smtplib

def process(orchestrator_connection: OrchestratorConnection) -> None:
    
    def Email(Modtagermail, file_name, file_path):
        SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
        SMTP_PORT = 25
        SCREENSHOT_SENDER = "aktbob@aarhus.dk"
        subject = "Indtastningsgrundlag i forhold til SD Løn"

        html = """
        <html>
        <body>
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
        msg["Cc"] = orchestrator_connection.get_constant('balas').value
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
        {"RunName": "SD Forfaldne faktura", "UploadMappe": "SP"},
        {"RunName": "SD Stamdatatabel", "UploadMappe": "SP"},
    ]

    for run in runs:
        if run["RunName"] == "SD løn udtræk":
            sap_running = initialize_sap(orchestrator_connection)
            if not sap_running:
                raise Exception("SAP failed to launch successfully")
            else:
                print("SAP is running and ready.")
            watcher = start_popup_watcher(interval= 0.3)
            try:
                print("▶ Starter SD løn udtræk")
                SDLonUdtrak(orchestrator_connection)
                
            finally:
                watcher.stop()
            Outfile, Name = InputToTemplate()

            upload_to_sharepoint(Client, Outfile, parent_folder_url, site_url_str=sharepoint_site_url)
            # Email(orchestrator_connection.get_constant('balas').value, file_name= Name, file_path= Outfile)
            file_deleter(Outfile)
            file_deleter('export.xlsx')

        elif run["RunName"] == "SD Forfaldne faktura":
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

        elif run["RunName"] == "SD Stamdatatabel":
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
        else:
            break