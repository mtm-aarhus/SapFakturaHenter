"""
Lægger én krypteret kø-element pr. leder med timer til gode i en OpenOrchestrator-kø.
Mail-processen henter elementerne og bruger dekrypter_besked().
"""
import json
import uuid
from datetime import datetime, timezone

import pandas as pd
from cryptography.fernet import Fernet, InvalidToken

NOEGLE_CREDENTIAL = "TimerReminderNoegle"   
MAX_DATA_LAENGDE = 2000                     


def _hent_fernet(noegle: str) -> Fernet:
    if not noegle:
        raise RuntimeError("Ingen krypteringsnøgle fundet")
    return Fernet(noegle.encode("ascii"))


def byg_beskeder(resultat: pd.DataFrame, noegle: str) -> tuple[list[str], list[str]]:
    """Returnerer (references, data) klar til bulk_create_queue_elements."""
    fernet = _hent_fernet(noegle)

    gyldige = resultat[
        resultat["leder_email"].notna()
        & resultat["leder_brugernavn"].notna()
        & (resultat["ikke_reg_timer"] > 0)
    ]
    sprunget_over = len(resultat) - len(gyldige)
    if sprunget_over:
        print(f"⚠  {sprunget_over} rækker uden leder, mail eller timer lægges ikke i køen")

    nu = datetime.now(timezone.utc)
    references, data = [], []
    for række in gyldige.itertuples(index=False):
        besked_id = str(uuid.uuid4())
        payload = {
            "id": besked_id,
            "oprettet": nu.isoformat(),
            "email": række.leder_email,
            "navn": række.leder_kaldenavn,
            "timer": round(float(række.ikke_reg_timer), 2),
        }
        token = fernet.encrypt(json.dumps(payload, ensure_ascii=False).encode("utf-8")).decode("ascii")
        if len(token) > MAX_DATA_LAENGDE:
            raise ValueError(f"Kø-data er for lang ({len(token)} tegn) for {besked_id}")

        # Reference står ukrypteret i Orchestrator, så den indeholder ingen personoplysninger
        references.append(f"TimerReminder_{nu:%Y%m%d}_{besked_id[:8]}")
        data.append(token)

    return references, data


def laeg_i_koe(orchestrator_connection, resultat: pd.DataFrame, koe_navn: str) -> int:
    """Krypterer og lægger beskederne i køen. Returnerer antal elementer."""
    noegle = orchestrator_connection.get_credential(NOEGLE_CREDENTIAL).password
    references, data = byg_beskeder(resultat, noegle)

    if not data:
        print("Ingen ledere med timer til gode, intet lagt i køen")
        return 0

    orchestrator_connection.bulk_create_queue_elements(
        koe_navn,
        references=tuple(references),
        data=tuple(data),
        created_by="MTMIkkeGodkendteTimer",
    )
    print(f"✅ {len(data)} kø-elementer lagt i '{koe_navn}'")
    return len(data)


def dekrypter_besked(besked: str, noegle: str, max_alder_sekunder: int | None = None) -> dict:
    """Bruges af mail-processen. max_alder_sekunder afviser gamle beskeder."""
    fernet = _hent_fernet(noegle)
    try:
        indhold = fernet.decrypt(besked.encode("ascii"), ttl=max_alder_sekunder)
    except InvalidToken:
        raise ValueError("Beskeden kunne ikke dekrypteres (forkert nøgle, ændret eller for gammel)")
    return json.loads(indhold.decode("utf-8"))