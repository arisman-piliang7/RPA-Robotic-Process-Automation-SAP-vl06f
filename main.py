"""
main.py
-------
Entry point for the RPA VL06F automatic report downloader.

Usage
-----
Run once::

    python main.py

Run on a recurring schedule (as defined in config.ini [SCHEDULE])::

    python main.py --schedule

Use a custom config file::

    python main.py --config path/to/my_config.ini

Print help::

    python main.py --help
"""

import argparse
import configparser
import logging
import sys

from utils.logger import setup_logger
from utils.file_utils import build_output_path


def load_config(config_path: str = "config.ini") -> configparser.ConfigParser:
    """
    Read *config_path* into a :class:`configparser.ConfigParser`.

    Raises
    ------
    FileNotFoundError
        If the config file does not exist.
    """
    cfg = configparser.ConfigParser()
    files_read = cfg.read(config_path, encoding="utf-8")
    if not files_read:
        raise FileNotFoundError(
            f"Config file not found: '{config_path}'.  "
            "Copy config.ini.example to config.ini and fill in your settings."
        )
    return cfg


def run_once(config: configparser.ConfigParser) -> str:
    """
    Execute one VL06F download run and return the output file path.

    Parameters
    ----------
    config:
        Loaded application configuration.

    Returns
    -------
    str
        Path of the saved report file.
    """
    from sap.connector import SAPConnector
    from sap.vl06f import VL06FDownloader

    logger = logging.getLogger("rpa_vl06f.main")

    output_path = build_output_path(
        output_dir_template=config["OUTPUT"].get("output_dir",  "output/{YYYY}/{MM}"),
        file_name_template=config["OUTPUT"].get("file_name",
                                                "VL06F_Report_{YYYY}{MM}{DD}_{HH}{MIN}{SS}.xlsx"),
    )

    logger.info("Output path: %s", output_path)

    connector = SAPConnector(config)
    try:
        session = connector.connect()
        downloader = VL06FDownloader(session, config)
        downloader.run(output_path)
    finally:
        connector.disconnect()

    return output_path


def main(argv=None) -> int:
    """
    Parse CLI arguments, set up logging, and dispatch to the appropriate
    execution mode.

    Parameters
    ----------
    argv:
        Argument list (defaults to :data:`sys.argv`).

    Returns
    -------
    int
        Exit code (0 = success, 1 = error).
    """
    parser = argparse.ArgumentParser(
        description="RPA – Automatic VL06F Report Downloader for SAP",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--config",
        default="config.ini",
        metavar="FILE",
        help="Path to the INI configuration file.",
    )
    parser.add_argument(
        "--schedule",
        action="store_true",
        help="Run on the recurring schedule defined in [SCHEDULE] section.",
    )
    args = parser.parse_args(argv)

    try:
        config = load_config(args.config)
    except FileNotFoundError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    log_cfg = config["LOGGING"] if "LOGGING" in config else {}
    logger = setup_logger(
        log_dir=log_cfg.get("log_dir", "logs"),
        log_level=log_cfg.get("log_level", "INFO"),
        backup_count=int(log_cfg.get("backup_count", 7)),
    )

    logger.info("=== RPA VL06F Report Downloader started ===")

    if args.schedule:
        from scheduler import VL06FScheduler
        scheduler = VL06FScheduler(
            download_func=lambda: run_once(config),
            config=config,
        )
        scheduler.start()
    else:
        try:
            output_path = run_once(config)
            logger.info("Download finished.  File: %s", output_path)
        except Exception as exc:
            logger.error("Download failed: %s", exc, exc_info=True)
            return 1

    logger.info("=== RPA VL06F Report Downloader finished ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
