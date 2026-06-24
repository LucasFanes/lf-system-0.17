import logging
from pathlib import Path

# Automatic log directory setup.
LOGS_FOLDER = Path.home() / "SISTEMA_LF" / "Logs"
LOGS_FOLDER.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    filename=LOGS_FOLDER / "Logs.txt",
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s",
    encoding="utf-8",
)

logging.info("================ LF SYSTEM STARTED ================")
