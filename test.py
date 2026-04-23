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
from scripts import SDLonUdtrak, InputToTemplate, SDForfaldneFaktura, SDStamdataTabel, MTMIkkeGodkendteTimer, SDAfstemning
from sap_popup_utils import start_popup_watcher
import os, time, shutil, tempfile, mimetypes
from email.message import EmailMessage
import smtplib



# Opsæt connection til Orchestrator
orchestrator_connection = OrchestratorConnection(
    "SapProcess",
    os.getenv('OpenOrchestratorSQL'),
    os.getenv('OpenOrchestratorKey'),
    None)
