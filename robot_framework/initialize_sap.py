# initialize_sap.py

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

import random
import string
import time
import psutil
import os
import win32com.client

from sap_popup_utils import sap_with_popup_guard, wait_for_main_sap_window
# hvis du vil type-annotere:
# from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection


def dismiss_until_easy_access(timeout: int = 30) -> bool:
    """
    Vent til SAP-session er klar og navigér til 'SAP Easy Access'.
    KØR DENNE INDE I:  with sap_with_popup_guard(...):
    """
    start_time = time.time()
    session = None

    # Find en session
    while time.time() - start_time < timeout:
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            application = sap_gui_auto.GetScriptingEngine
            if application.Children.Count > 0:
                connection = application.Children(0)
                if connection.Children.Count > 0:
                    session = connection.Children(0)
                    print("SAP session is ready.")
                    break
        except Exception:
            pass
        time.sleep(0.5)

    if not session:
        raise TimeoutError("SAP session not available within timeout.")

    # Navigér til SAP Easy Access
    print("Checking for SAP Easy Access...")
    while time.time() - start_time < timeout:
        try:
            wnd0 = session.findById("wnd[0]")
            title = (wnd0.Text or "").strip()

            if title.startswith("SAP Easy Access"):
                print("SAP Easy Access screen reached.")
                return True

            # Prøv 'Tilbage'-knap eller ESC
            try:
                session.ActiveWindow.FindById("tbar[0]/btn[0]").Press()
                print(f"Dismissed '{session.ActiveWindow.Text.strip()}' via tbar[0]/btn[0]")
            except Exception:
                try:
                    session.ActiveWindow.SendVKey(12)  # ESC
                except Exception:
                    pass

        except Exception:
            pass

        time.sleep(0.5)

    raise TimeoutError("SAP Easy Access screen not reached within timeout.")


def download_sap(driver: webdriver.Chrome, downloads_folder: str, orchestrator_connection, parent_tab):
    """Klik på fanen/elementet og vent på at en .sap-fil lander i mappen."""
    before = set(os.listdir(downloads_folder))
    driver.execute_script("arguments[0].click();", parent_tab)

    start_time = time.time()
    timeout = 10

    while time.time() - start_time < timeout:
        time.sleep(0.25)
        after = set(os.listdir(downloads_folder))
        new_files = after - before
        if new_files:
            for file in new_files:
                if file.endswith(".sap"):
                    full_path = os.path.join(downloads_folder, file)
                    orchestrator_connection.log_info(f"Found SAP file: {file}")
                    return full_path
    raise TimeoutError("SAP file not downloaded.")


def initialize_sap(orchestrator_connection):
    """Logger ind i Opus, downloader .sap, starter SAP Logon og går til Easy Access med tidlig popup-guard."""
    # Opus bruger
    OpusLogin = orchestrator_connection.get_credential("OpusBruger")
    OpusUser = OpusLogin.username
    OpusPassword = OpusLogin.password

    # Robotpassword (hvis password-udskiftning er nødvendig)
    RobotCredential = orchestrator_connection.get_credential("Robot365User")
    RobotUsername = RobotCredential.username
    RobotPassword = RobotCredential.password

    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

    # Chrome setup
    chrome_options = Options()
    chrome_options.add_argument('--remote-debugging-pipe')
    # chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": downloads_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
    })

    driver = webdriver.Chrome(options=chrome_options)
    driver.get(orchestrator_connection.get_constant("OpusAdgangUrl").value)
    orchestrator_connection.log_info("Navigating to Opus login page")

    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "logonuidfield")))
    driver.find_element(By.ID, "logonuidfield").send_keys(OpusUser)
    driver.find_element(By.ID, "logonpassfield").send_keys(OpusPassword)
    driver.find_element(By.ID, "buttonLogon").click()

    orchestrator_connection.log_info("Logged in to Opus portal successfully")

    WebDriverWait(driver, 60).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    tab_label_xpath = "//div[contains(@class, 'TabText_SmallTabs') and contains(text(), 'Mine Genveje')]"

    try:
        tab_label = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, tab_label_xpath))
        )
        parent_tab = tab_label.find_element(By.XPATH, "./ancestor::div[contains(@id, 'tabIndex')]")

    except:
        orchestrator_connection.log_info('Trying to find change button')
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "changeButton")))
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.ID, "changeButton")))

        # Generér nyt password iht. politikken
        lower = string.ascii_lowercase
        upper = string.ascii_uppercase
        digits = string.digits
        special = "!@#&%"

        password_chars = []
        password_chars += random.choices(lower, k=2)
        password_chars += random.choices(upper, k=2)
        password_chars += random.choices(digits, k=4)
        password_chars += random.choices(special, k=2)

        random.shuffle(password_chars)
        password = ''.join(password_chars)

        driver.find_element(By.ID, "inputUsername").send_keys(OpusPassword)
        driver.find_element(By.NAME, "j_sap_password").send_keys(password)
        driver.find_element(By.NAME, "j_sap_again").send_keys(password)
        driver.find_element(By.ID, "changeButton").click()

        orchestrator_connection.update_credential('OpusBruger', OpusUser, password)
        orchestrator_connection.log_info('Password changed and credential updated')

        time.sleep(2)
        driver.get(orchestrator_connection.get_constant("OpusAdgangUrl").value)

        tab_label = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, tab_label_xpath))
        )
        parent_tab = tab_label.find_element(By.XPATH, "./ancestor::div[contains(@id, 'tabIndex')]")

    filepath = download_sap(driver, downloads_folder, orchestrator_connection, parent_tab)
    driver.quit()

    # Start SAP Logon via .sap
    os.startfile(filepath)

    # TIDLIG popup-guard dækker både proces-vent, hovedvindue og overgang til Easy Access
    with sap_with_popup_guard(interval=1.0):

        # Vent på at SAP-processen starter
        start_time = time.time()
        while time.time() - start_time < 30:
            if any('saplogon' in (p.info['name'] or '').lower() for p in psutil.process_iter(['name'])):
                break
            time.sleep(0.5)

        # Vent på hovedvinduet (fra sap_popup_utils)
        wait_for_main_sap_window()

        # Navigér til Easy Access
        dismiss_until_easy_access(30)

    return True
