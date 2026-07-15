"""
scheduler.py
------------
APScheduler-based wrapper that runs the VL06F download job on a cron
schedule defined in ``config.ini``.

Usage
-----
Called automatically by ``main.py --schedule``.  You can also import
:class:`VL06FScheduler` directly and call :meth:`VL06FScheduler.start`.
"""

import logging

from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.cron import CronTrigger

logger = logging.getLogger("rpa_vl06f.scheduler")


class VL06FScheduler:
    """
    Wraps an APScheduler :class:`BlockingScheduler` and registers
    the VL06F download job using cron fields from the ``[SCHEDULE]``
    config section.

    Parameters
    ----------
    download_func : callable
        Zero-argument callable that performs one VL06F download run.
        Typically this is :func:`main.run_once`.
    config : configparser.ConfigParser
        Full application configuration.
    """

    def __init__(self, download_func, config) -> None:
        self._func   = download_func
        self._cfg    = config["SCHEDULE"]
        self._scheduler = BlockingScheduler(
            timezone=self._cfg.get("timezone", "Asia/Jakarta")
        )

    def start(self) -> None:
        """
        Register the job and start the scheduler (blocks until stopped).

        The job is also executed immediately on startup so that the first
        run does not have to wait for the next scheduled time.
        """
        trigger = CronTrigger(
            day_of_week=self._cfg.get("cron_day_of_week", "mon-fri"),
            hour=int(self._cfg.get("cron_hour",   7)),
            minute=int(self._cfg.get("cron_minute", 0)),
            timezone=self._cfg.get("timezone", "Asia/Jakarta"),
        )

        self._scheduler.add_job(
            func=self._run_job,
            trigger=trigger,
            id="vl06f_download",
            name="VL06F Report Download",
            replace_existing=True,
        )

        logger.info(
            "Scheduler started.  Next run: %s",
            self._scheduler.get_jobs()[0].next_run_time,
        )

        # Run immediately on startup
        logger.info("Running initial download before entering schedule loop …")
        self._run_job()

        try:
            self._scheduler.start()
        except (KeyboardInterrupt, SystemExit):
            logger.info("Scheduler stopped by user.")

    # ------------------------------------------------------------------

    def _run_job(self) -> None:
        """Execute one download attempt, logging any errors without crashing."""
        logger.info("Scheduled job: starting VL06F download …")
        try:
            self._func()
            logger.info("Scheduled job: download completed successfully.")
        except Exception as exc:
            logger.error("Scheduled job failed: %s", exc, exc_info=True)
