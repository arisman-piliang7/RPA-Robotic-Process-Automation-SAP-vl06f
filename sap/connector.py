"""
sap/connector.py
----------------
Manages the SAP GUI COM connection lifecycle.

Requirements
------------
* SAP GUI for Windows must be installed on the host machine.
* SAP GUI scripting must be enabled:
    SAP GUI → Options → Accessibility & Scripting → Scripting → Enable.
* pywin32 must be installed:  pip install pywin32
"""

import logging
import time

logger = logging.getLogger("rpa_vl06f.connector")


class SAPConnector:
    """
    Opens and owns a single SAP GUI session.

    Usage::

        connector = SAPConnector(config)
        session   = connector.connect()
        # ... use session ...
        connector.disconnect()

    Parameters
    ----------
    config : configparser.ConfigParser
        Full application configuration.  The ``[SAP]`` section is used.
    """

    def __init__(self, config) -> None:
        self._cfg          = config["SAP"]
        self._application  = None
        self._connection   = None
        self.session       = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def connect(self):
        """
        Attach to the running SAP GUI Scripting Engine and open a new
        connection/session.

        Returns
        -------
        session
            The active ``GuiSession`` COM object.

        Raises
        ------
        RuntimeError
            If SAP GUI is not running or scripting is disabled.
        """
        try:
            import win32com.client  # noqa: PLC0415 – Windows-only import
        except ImportError as exc:
            raise RuntimeError(
                "pywin32 is not installed.  "
                "Run: pip install pywin32"
            ) from exc

        logger.info("Connecting to SAP GUI scripting engine …")
        try:
            sap_gui = win32com.client.GetActiveObject("SapGui.ScriptingCtrl.1")
        except Exception as exc:
            raise RuntimeError(
                "Cannot reach SAP GUI scripting engine.  "
                "Make sure SAP GUI is running and scripting is enabled."
            ) from exc

        self._application = sap_gui.GetScriptingEngine

        system = self._cfg.get("system_description", "")
        logger.info("Opening SAP connection to '%s' …", system)
        self._connection = self._application.OpenConnection(system, True)

        self.session = self._connection.Children(0)

        # Log on
        self._logon()
        logger.info("SAP session ready.")
        return self.session

    def disconnect(self) -> None:
        """Close the SAP connection gracefully."""
        if self._connection is not None:
            try:
                self._connection.CloseSession(self.session.Id)
            except Exception:
                pass
            self._connection = None
            self.session = None
            logger.info("SAP connection closed.")

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _logon(self) -> None:
        """Fill in logon credentials on the SAP initial screen."""
        timeout = int(self._cfg.get("screen_wait_timeout", 30))
        self._wait_for_screen("SAPLOGIN", timeout)

        session = self.session
        client   = self._cfg.get("client",   "100")
        username = self._cfg.get("username", "")
        password = self._cfg.get("password", "")
        language = self._cfg.get("language", "EN")

        logger.debug("Filling logon screen (client=%s, user=%s) …", client, username)

        session.FindById("wnd[0]/usr/txtRSYST-MANDT").Text = client
        session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = username
        session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = password
        session.FindById("wnd[0]/usr/txtRSYST-LANGU").Text = language

        session.FindById("wnd[0]").SendVKey(0)  # Enter

        self._wait_for_ready(timeout)
        logger.debug("Logon complete.")

    def _wait_for_screen(self, expected_program: str, timeout: int) -> None:
        """Poll until the session shows *expected_program* or timeout."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                info = self.session.Info
                if info.Program == expected_program:
                    return
            except Exception:
                pass
            time.sleep(0.5)

    def _wait_for_ready(self, timeout: int) -> None:
        """Poll until the SAP session is no longer busy."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                if not self.session.Busy:
                    return
            except Exception:
                pass
            time.sleep(0.3)
