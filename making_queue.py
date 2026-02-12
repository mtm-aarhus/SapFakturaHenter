from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import os
#Connect to orchestrator
orchestrator_connection = OrchestratorConnection("Opus bookmark performer", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)

orchestrator_connection.create_queue_element("OpusBookmarkQueue")