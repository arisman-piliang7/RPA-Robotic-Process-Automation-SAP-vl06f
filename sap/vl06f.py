"""
sap/vl06f.py
------------
Automates SAP transaction VL06F (Outbound Deliveries – Goods Issues).

Steps performed
---------------
1. Start transaction VL06F via the command field.
2. Clear the selection screen and apply configured parameters
   (shipping point, date range).
3. Execute the report (F8).
4. Export the result list to a local file (spreadsheet or text).
5. Return the path to the saved file.
"""

import logging
import os
import time
from datetime import date, datetime, timedelta

logger = logging.getLogger("rpa_vl06f.vl06f")

# SAP virtual-key constants used in this module
VK_ENTER     = 0
VK_F8        = 8   # Execute
VK_F3        = 3   # Back


class VL06FDownloader:
    """
    Drives the SAP VL06F transaction and saves the report.

    Parameters
    ----------
    session : GuiSession
        Active SAP GUI session returned by :class:`~sap.connector.SAPConnector`.
    config : configparser.ConfigParser
        Full application configuration.
    """

    def __init__(self, session, config) -> None:
        self._session = session
        self._vl06f_cfg = config["VL06F"]
        self._output_cfg = config["OUTPUT"]
        self._timeout = int(config["SAP"].get("screen_wait_timeout", 30))

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def run(self, output_path: str) -> str:
        """
        Execute VL06F and save the result to *output_path*.

        Parameters
        ----------
        output_path:
            Full file path (including extension) where the report will be
            saved.  The parent directory must exist before calling this
            method.

        Returns
        -------
        str
            The resolved *output_path* that was written.
        """
        logger.info("Starting VL06F download …")

        self._start_transaction()
        self._fill_selection_screen()
        self._execute_report()
        self._export_report(output_path)

        logger.info("Report saved to: %s", output_path)
        return output_path

    # ------------------------------------------------------------------
    # Internal steps
    # ------------------------------------------------------------------

    def _start_transaction(self) -> None:
        """Navigate to VL06F using the transaction command field."""
        logger.debug("Starting transaction VL06F …")
        self._session.StartTransaction("VL06F")
        self._wait_for_ready()

    def _fill_selection_screen(self) -> None:
        """Populate the VL06F selection screen with configured parameters."""
        logger.debug("Filling VL06F selection screen …")
        session = self._session

        # --- Shipping point -------------------------------------------
        shipping_points = [
            sp.strip()
            for sp in self._vl06f_cfg.get("shipping_point", "").split(",")
            if sp.strip()
        ]
        if shipping_points:
            try:
                field = session.FindById(
                    "wnd[0]/usr/ctxtLIKP-VSTEL"
                )
                field.Text = shipping_points[0]
                logger.debug("Shipping point set to: %s", shipping_points[0])
            except Exception:
                logger.warning(
                    "Could not set shipping point field; "
                    "it may have a different control ID on this system."
                )

        # --- Date range -----------------------------------------------
        date_from, date_to = self._resolve_date_range()
        date_from_str = date_from.strftime("%d.%m.%Y")
        date_to_str   = date_to.strftime("%d.%m.%Y")

        try:
            session.FindById(
                "wnd[0]/usr/ctxtLIKP-WADAT_IST-LOW"
            ).Text = date_from_str
            session.FindById(
                "wnd[0]/usr/ctxtLIKP-WADAT_IST-HIGH"
            ).Text = date_to_str
            logger.debug(
                "Date range: %s – %s", date_from_str, date_to_str
            )
        except Exception:
            logger.warning(
                "Could not set date range fields; "
                "control IDs may differ on this SAP system."
            )

        # --- Layout variant -------------------------------------------
        variant = self._vl06f_cfg.get("layout_variant", "").strip()
        if variant:
            try:
                session.FindById(
                    "wnd[0]/usr/txtLIKP-VBELN"
                ).Text = variant
            except Exception:
                logger.debug("No layout variant field found; skipping.")

    def _execute_report(self) -> None:
        """Press F8 to execute the report and wait until the list appears."""
        logger.debug("Executing VL06F report (F8) …")
        self._session.FindById("wnd[0]").SendVKey(VK_F8)
        self._wait_for_ready()
        logger.debug("Report executed.")

    def _export_report(self, output_path: str) -> None:
        """
        Export the report list to *output_path* via
        System → List → Save → Local File.
        """
        logger.debug("Exporting report to %s …", output_path)
        ext = os.path.splitext(output_path)[1].lower()

        session = self._session

        # Open the export dialog via menu path
        try:
            session.FindById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select()
        except Exception:
            # Fall back to keyboard shortcut (Ctrl+Shift+F9 in some versions)
            logger.debug(
                "Menu path for export not found; "
                "trying keyboard shortcut …"
            )
            session.FindById("wnd[0]").SendVKey(43)

        self._wait_for_ready()

        # Choose the format in the dialog
        if ext == ".xlsx" or ext == ".xls":
            format_id = "btn[1]"   # Spreadsheet
        else:
            format_id = "btn[0]"   # Unconverted (tab-separated text)

        try:
            popup = session.FindById("wnd[1]")
            popup.FindById(f"usr/{format_id}").Select()
            popup.FindById("usr/btnBUTTON_1").Press()     # Continue/Generate
        except Exception as exc:
            logger.error("Export dialog interaction failed: %s", exc)
            raise

        self._wait_for_ready()

        # Type the output file path in the file-name dialog
        try:
            file_dialog = session.FindById("wnd[1]")
            filename_field = file_dialog.FindById("usr/ctxtDY_FILENAME")
            filename_field.Text = os.path.abspath(output_path)
            file_dialog.FindById("usr/btnBUTTON_1").Press()  # Save
        except Exception as exc:
            logger.error("File save dialog interaction failed: %s", exc)
            raise

        self._wait_for_ready()
        logger.debug("Export dialog closed.")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _resolve_date_range(self) -> tuple[date, date]:
        """
        Return a ``(date_from, date_to)`` tuple based on the configuration.

        Supported modes
        ---------------
        ``today``       – today's date for both from and to.
        ``yesterday``   – yesterday for both from and to.
        ``last_n_days`` – from (today − N + 1) to today.
        ``custom``      – explicit ``date_from`` / ``date_to`` values.
        """
        mode  = self._vl06f_cfg.get("date_range_mode", "today").lower().strip()
        today = date.today()

        if mode == "today":
            return today, today

        if mode == "yesterday":
            yesterday = today - timedelta(days=1)
            return yesterday, yesterday

        if mode == "last_n_days":
            n = int(self._vl06f_cfg.get("last_n_days", 7))
            return today - timedelta(days=n - 1), today

        if mode == "custom":
            fmt = "%d.%m.%Y"
            d_from = datetime.strptime(
                self._vl06f_cfg.get("date_from", "01.01.2024"), fmt
            ).date()
            d_to = datetime.strptime(
                self._vl06f_cfg.get("date_to", "31.12.2024"), fmt
            ).date()
            return d_from, d_to

        logger.warning(
            "Unknown date_range_mode '%s'; defaulting to today.", mode
        )
        return today, today

    def _wait_for_ready(self) -> None:
        """Block until the SAP session is no longer busy."""
        deadline = time.time() + self._timeout
        while time.time() < deadline:
            try:
                if not self._session.Busy:
                    return
            except Exception:
                pass
            time.sleep(0.3)
