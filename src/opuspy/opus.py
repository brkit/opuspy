import subprocess
import time
import winreg

from loguru import logger
import win32com.client  # pywin32

from brkrpautils import (
    get_credentials,
    backup_old_password,
    generate_new_password,
    save_new_password,
)

from contextlib import contextmanager

import pythoncom


def say_hello_from_opuspy():
    logger.info("Hello from opuspy")


@contextmanager
def sap_connection(timeout=30, interval=1, auto_close=False, force_close=False):
    """
    Attempt to bind to an existing SAP GUI session within a given timeout.
    Yields the session once available.
    If auto_close is True, log off / close when the 'with' block exits.
    If force_close is True, kill the SAP processes if polite log-off fails.
    """
    pythoncom.CoInitialize()

    start_time = time.time()
    session = None

    while True:
        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
            scripting_engine  = sap_gui.GetScriptingEngine
            if scripting_engine and scripting_engine.Children.Count:
                connection = scripting_engine.Children(0)      # first connection
                if connection.Children.Count:
                    session = connection.Children(0) # first session
                    break
        except pythoncom.com_error:
            pass

        if (time.time() - start_time) > timeout:
            pythoncom.CoUninitialize()
            raise RuntimeError("Timed out waiting for SAP GUI readiness.")

        time.sleep(interval)

    try:
        yield session
    finally:
        if auto_close:
            try:
                _sap_logoff(session, force=force_close)
            except Exception:
                pass        # don't mask upstream exceptions
        session = None
        pythoncom.CoUninitialize()

def _sap_logoff(session, force=False):
    """
    Try to exit SAP politely through its scripting interface.
    When force=True, kill the processes if polite log-off fails.
    """
    if session is None:
        return

    try:
        # put /nex in OK-code box -> "log off all sessions"
        okcd = session.findById("wnd[0]/tbar[0]/okcd", False)
        if okcd:
            okcd.text = "/nex"
            session.findById("wnd[0]").sendVKey(0)           # press <Enter>

        # confirm the "Log off" dialog if it appears
        try:
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()  # “Yes”
        except Exception:
            pass

        if not force:
            return  # successfully closed connection/session window (or at least we asked politely)

    except Exception:         # scripting call failed / window hung
        if not force:
            return

    # ── If we arrive here we either asked for force=True or polite close failed ──
    for exe in ("saplogon.exe", "saplgpad.exe", "sapgui.exe"):
        subprocess.run(["taskkill", "/IM", exe, "/F"],
                       capture_output=True, check=False)

def start_opus(pam_path, user, sapshcut_path):
    """
    Start SAP session with SAPSHCUT.exe
    :param pam_path: str, path to PAM file
    :param user: str, user to log in as
    :param sapshcut_path: str, path to SAPSHCUT.exe
    """

    # Unpack credentials
    username, password = get_credentials(pam_path, user, fagsystem="opus")

    if not username or not password:
        logger.error("Failed to retrieve credentials.")
        return None

    command_args = [
        str(sapshcut_path),
        "-system=P02",
        "-client=400",
        f"-user={username}",
        f"-pw={password}",
    ]

    subprocess.run(command_args, check=False)  # noqa: S603

    with sap_connection(auto_close=False, force_close=False) as session:
        # Check if SAP with ID /app/con[0]/ses[0]/wnd[1]/usr is open to determine if password reset prompt is present
        element_id = "/app/con[0]/ses[0]/wnd[1]/usr/lblRSYST-NCODE_TEXT"
        try:
            element = session.findById(element_id)
        except Exception as e:
            return  # password prompt not present, continuing as normal

        if element.text == "Nyt password":
            try:
                logger.info("Detected password reset prompt in SAP.")

                backup_old_password(pam_path=pam_path, user=user)
                new_password = generate_new_password(17)
                save_new_password(
                    new_password=new_password,
                    pam_path=pam_path,
                    user=user,
                    fagsystem="opus",
                )

                # Write new password to SAP
                session.findById(
                    "/app/con[0]/ses[0]/wnd[1]/usr/pwdRSYST-NCODE"
                ).text = new_password
                session.findById(
                    "/app/con[0]/ses[0]/wnd[1]/usr/pwdRSYST-NCOD2"
                ).text = new_password

                # Press OK
                session.findById("/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]").press()
                logger.info("Password updated successfully in SAP.")
                time.sleep(1)
            except Exception as e:
                logger.error(f"Error while trying to change password {element_id}: {e}")
                raise e

        else:
            logger.info("element_id found, but text did not match 'Nyt password'.")
            raise RuntimeError(
                "element_id found, but text did not match 'Nyt password'."
            )

def is_sap_scripting_allowed():
    """
    TRUE if scripting is allowed, FALSE if not.
    """
    try:
        # Define the registry path
        registry_path = (
            r"SOFTWARE\WOW6432Node\SAP\SAPGUI Front\SAP Frontend Server\Security"
        )
        key_name = "UserScripting"

        # Open the registry key
        reg_key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE, registry_path, 0, winreg.KEY_READ
        )

        # Get the value of the UserScripting key
        value, regtype = winreg.QueryValueEx(reg_key, key_name)

        # if value is 1, scripting is allowed
        if value == 1:
            return True
        else:
            return False

        # Close the registry key
        winreg.CloseKey(reg_key)
    except FileNotFoundError:
        print("The specified registry key or value does not exist.")
    except PermissionError:
        print("Permission denied. Please run the script as an administrator.")
    except Exception as e:
        print(f"An error occurred: {e}")

def set_sap_scripting_to_allowed():
    """
    Set SAP scripting to allowed by changing the registry key.
    """
    try:
        # Define the registry path
        registry_path = (
            r"SOFTWARE\WOW6432Node\SAP\SAPGUI Front\SAP Frontend Server\Security"
        )
        key_name = "UserScripting"

        # Open the registry key with write access
        reg_key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE, registry_path, 0, winreg.KEY_WRITE
        )

        # Set the value of the UserScripting key to 1 (allowed)
        winreg.SetValueEx(reg_key, key_name, 0, winreg.REG_DWORD, 1)

        # Close the registry key
        winreg.CloseKey(reg_key)
        print("SAP scripting has been set to allowed.")
    except FileNotFoundError:
        print("The specified registry key or value does not exist.")
    except PermissionError:
        print("Permission denied. Please run the script as an administrator.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    if is_sap_scripting_allowed():
        start_opus()
    else:
        print("SAP scripting is not allowed. Please ask admin to enable scripting in registry")
