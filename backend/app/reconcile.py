"""Periodic recovery of transactions committed before an RQ enqueue failure."""

import logging
import signal
from threading import Event

from .jobs import reconcile_pending

logger = logging.getLogger(__name__)
stopped = Event()


def main() -> None:
    logging.basicConfig(level=logging.INFO)

    def stop(*_):
        stopped.set()

    signal.signal(signal.SIGTERM, stop)
    signal.signal(signal.SIGINT, stop)
    while not stopped.is_set():
        try:
            reconcile_pending()
        except Exception:
            logger.exception("Pending job reconciliation failed")
        stopped.wait(30)


if __name__ == "__main__":
    main()

