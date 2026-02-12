import time
import win32gui
import win32con
import threading
from contextlib import contextmanager
sap_main_hwnd = None  # Registrerer SAP-hovedvindue
# top: tilføj
from typing import Optional
import threading

class PopupWatcher:
    def __init__(self, interval: float = 1.0):
        self._stop = threading.Event()
        self._interval = interval
        self._thr = threading.Thread(target=self._run, name="SAPPopupWatcher", daemon=False)

    def _run(self):
        while not self._stop.is_set():
            try:
                close_all_sap_popups(timeout=3)
            except Exception as e:
                print(f"⚠️ Popup-watcher fejl: {e}")
            self._stop.wait(self._interval)  # kan afbrydes straks

    def start(self):
        self._thr.start()
        return self

    def stop(self, join_timeout: float = 10.0):
        self._stop.set()
        self._thr.join(join_timeout)

@contextmanager
def sap_with_popup_guard(interval: float = 1.0):
    watcher = PopupWatcher(interval=interval).start()
    try:
        yield
    finally:
        watcher.stop(join_timeout=10.0)

def start_popup_watcher(interval: float = 1.0):
    """Hvis du VIL starte den manuelt, så returnér watcher-objektet og HUSK stop()."""
    return PopupWatcher(interval=interval).start()


def diagnose_sap_popup(timeout=10):
    """
    Logger detaljer om alle SAP-relaterede popup-vinduer.
    Bruges til fejlsøgning.
    """
    print("▶ Starter SAP-popup diagnosticering...")
    end_time = time.time() + timeout

    def enum_callback(hwnd, results):
        title = win32gui.GetWindowText(hwnd)
        if "SAP" in title:
            results.append(hwnd)

    while time.time() < end_time:
        found = []
        win32gui.EnumWindows(enum_callback, found)

        for hwnd in found:
            try:
                win32gui.SetForegroundWindow(hwnd)
            except:
                pass

            def enum_children_callback(child_hwnd, _):
                c_type = win32gui.GetClassName(child_hwnd)
                c_text = win32gui.GetWindowText(child_hwnd)
            win32gui.EnumChildWindows(hwnd, enum_children_callback, None)

        if found:
            return

        time.sleep(0.5)
def close_all_sap_popups(timeout=10):
    """
    Lukker SAP-popups. Tolerant heuristik:
      - Ignorerer hovedvinduer
      - Accepterer EITHER kendte nøgleord OR tydelig OK/JA/Tillad-knap
    """
    print("▶ Søger og lukker kendte SAP-popups...")
    deadline = time.time() + timeout
    last_action_time = time.time()

    # Brug brede nøgleord – dansk/engelsk
    KEYWORDS = {
        "et script forsøger", "gui security", "sikkerhed", "certificate", "certifikat",
        "tillad adgang", "adgang", "gem som", "filtype", "loginforsøg", "login attempt",
        "open", "åbn", "run", "kør"
    }

    # Knap-tekster vi vil trykke på (lowercased)
    OK_BUTTONS = {"ok", "&ok", "ja", "&ja", "yes", "&yes", "tillad", "&tillad",
                  "tillad adgang", "allow", "allow access"}

    def is_known_popup(hwnd):
        title = win32gui.GetWindowText(hwnd).strip().lower()

        IGNORED_TITLE_FRAGMENTS = {
            "easy access",
            "sap gui for windows",
            "sap easy access",
            "p02_prod_kunder",
        }
        if any(frag in title for frag in IGNORED_TITLE_FRAGMENTS):
            return False

        found_keyword = False
        found_ok_like = False

        def scan_child(child_hwnd, _):
            nonlocal found_keyword, found_ok_like
            txt = win32gui.GetWindowText(child_hwnd).strip().lower()
            if any(k in txt for k in KEYWORDS):
                found_keyword = True
            if txt in OK_BUTTONS:
                found_ok_like = True

        try:
            win32gui.EnumChildWindows(hwnd, scan_child, None)
        except:
            return False

        # Mere tolerant: nøjes med ét af kriterierne
        return found_keyword or found_ok_like

    def try_close(hwnd):
        global sap_main_hwnd
        if hwnd == sap_main_hwnd:
            return False

        if not is_known_popup(hwnd):
            return False

        # 1) Prøv at trykke knappen (OK/JA/Tillad/Allow)
        try:
            def click_ok(child_hwnd, _):
                txt = win32gui.GetWindowText(child_hwnd).strip().lower()
                if txt in OK_BUTTONS:
                    try:
                        win32gui.SetForegroundWindow(hwnd)
                    except:
                        pass
                    time.sleep(0.2)
                    win32gui.SendMessage(child_hwnd, win32con.BM_CLICK, 0, 0)
                    raise StopIteration
            win32gui.EnumChildWindows(hwnd, click_ok, None)
        except StopIteration:
            return True

        # 2) Ellers luk vinduet (fallback)
        try:
            print(f"🔺 Lukker popup med WM_CLOSE: {win32gui.GetWindowText(hwnd)}")
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            return True
        except:
            return False

    while time.time() < deadline:
        windows = []

        def enum_windows(hwnd, _):
            if "SAP" in win32gui.GetWindowText(hwnd):
                windows.append(hwnd)

        win32gui.EnumWindows(enum_windows, None)

        any_closed = False
        for hwnd in windows:
            if try_close(hwnd):
                any_closed = True
                last_action_time = time.time()

        # Stop hvis der ikke er flere relevante popups i et par sekunder
        if not any_closed and time.time() - last_action_time > 2:
            print("✅ Ingen flere relevante SAP-popups.")
            return True

        time.sleep(0.3)

    print("❌ Timeout – popup-vinduer kunne ikke lukkes alle.")
    return False

def watch_and_dismiss_popup(timeout=15):
    """
    Søger efter SAP-popups og lukker dem via OK-knap eller kryds.
    Lukker kun relevante vinduer (f.eks. 'Et script forsøger' eller 'mislykkede loginforsøg').
    """
    print("▶ Starter SAP-popup watcher...")
    deadline = time.time() + timeout

    time.sleep(4)  # Giv popup tid til at åbne

    def try_dismiss(hwnd):
        title = win32gui.GetWindowText(hwnd).lower()
        if not title.startswith("sap"):
            return False

        class HitOK(Exception): pass

        try:
            def enum_child_callback(child_hwnd, _):
                class_name = win32gui.GetClassName(child_hwnd)
                text = win32gui.GetWindowText(child_hwnd).strip().lower()
                if class_name == "Button" and text in {"ok", "&ok"}:
                    print(f"✅ Klikker OK i '{win32gui.GetWindowText(hwnd)}'")
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.2)
                    win32gui.SendMessage(child_hwnd, win32con.BM_CLICK, 0, 0)
                    raise HitOK()

            win32gui.EnumChildWindows(hwnd, enum_child_callback, None)

        except HitOK:
            return True

        # Hvis ingen OK-knap, prøv med luk (krydset)
        try:
            print(f"🔺 Lukker vindue '{title}' med WM_CLOSE")
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            return True
        except Exception as e:
            print(f"❌ Fejl ved lukning af vindue: {e}")
            return False

    while time.time() < deadline:
        matches = []

        def enum_windows_callback(hwnd, _):
            title = win32gui.GetWindowText(hwnd)
            if "SAP" in title:
                matches.append(hwnd)

        win32gui.EnumWindows(enum_windows_callback, None)

        for hwnd in matches:
            if try_dismiss(hwnd):
                return

        time.sleep(0.3)

    print("❌ Timeout: Ingen SAP-popup blev håndteret.")
    diagnose_sap_popup(timeout=5)
    raise Exception("Kunne ikke lukke popup inden for timeout.")

def safe_sap_action(action_fn, retries=2):
    """
    Wrapper til SAP GUI-kommandokald, som håndterer kendte popups og prøver igen.
    """
    for attempt in range(retries):
        try:
            return action_fn()
        except Exception as e:
            print(f"⚠️ Fejl under SAP-handling: {e}")
            print("▶ Tjekker for og forsøger at lukke popup...")
            watch_and_dismiss_popup(timeout=5)
    raise Exception("❌ SAP-handling mislykkedes gentagne gange.")

def wait_for_main_sap_window(timeout=15):
    """
    Finder og gemmer hwnd for hovedvinduet 'SAP GUI for Windows ...'
    """
    global sap_main_hwnd
    print("⌛ Venter på SAP-hovedvindue...")
    deadline = time.time() + timeout

    while time.time() < deadline:
        def check(hwnd, result):
            title = win32gui.GetWindowText(hwnd).lower()
            if "sap gui for windows" in title:
                result.append(hwnd)

        found = []
        win32gui.EnumWindows(check, found)
        if found:
            sap_main_hwnd = found[0]
            print(f"📌 Hovedvindue registreret: '{win32gui.GetWindowText(sap_main_hwnd)}'")
            return

        time.sleep(0.5)

    raise TimeoutError("Kunne ikke finde SAP-hovedvindue i tide.")
