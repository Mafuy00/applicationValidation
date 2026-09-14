#!/usr/bin/env python
"""
Stores the SIP intranet credentials in Windows Credential Manager, so they
never have to be written into a file.

This is the recommended way to configure the upload automation on this
machine, because the repo sits inside a synced OneDrive folder - a plaintext
.env there would be copied to the cloud. Credential Manager encrypts the entry
against your Windows account instead.

Run from the upload_application_automation folder:

    python set_credentials.py            # prompt and store (password not echoed)
    python set_credentials.py --show     # what is configured (never prints the password)
    python set_credentials.py --clear    # remove the stored credentials
"""
import argparse
import getpass
import sys

from sip_uploader import config


def _keyring():
    try:
        import keyring
    except ImportError:
        print(
            "The 'keyring' package is required to use Windows Credential Manager.\n"
            "Install it with:  pip install -r requirements.txt"
        )
        sys.exit(1)
    return keyring


def _warn_if_env_shadows():
    """
    Environment variables win over the credential store, so a leftover .env
    would silently override what we just saved. Say so rather than let it
    confuse a later run.
    """
    env_user, env_password = config._credentials_from_env()
    if env_user and env_password:
        print(
            f"\n[warning] {config.USERNAME_ENV_VAR} / {config.PASSWORD_ENV_VAR} are also set "
            f"in the environment\n"
            f"          (probably from '{config.ENV_FILE_PATH}'), and those take precedence.\n"
            f"          Clear or delete that file for the stored credentials to be used."
        )


def show():
    keyring = _keyring()
    username = keyring.get_password(config.KEYRING_SERVICE_NAME, config.KEYRING_USERNAME_KEY)
    password = keyring.get_password(config.KEYRING_SERVICE_NAME, config.KEYRING_PASSWORD_KEY)

    print(f"Credential store : {keyring.get_keyring().__class__.__name__}")
    print(f"Service name     : {config.KEYRING_SERVICE_NAME}")
    print(f"Stored username  : {username or '(none)'}")
    print(f"Stored password  : {'(set)' if password else '(none)'}")
    print(f"\nA run right now would take credentials from: "
          f"{config.credentials_source() or '(nothing configured)'}")
    _warn_if_env_shadows()


def clear():
    keyring = _keyring()
    removed = 0
    for key in (config.KEYRING_USERNAME_KEY, config.KEYRING_PASSWORD_KEY):
        try:
            keyring.delete_password(config.KEYRING_SERVICE_NAME, key)
            removed += 1
        except Exception:
            pass  # nothing stored under that key
    print(f"Removed {removed} stored credential entr{'y' if removed == 1 else 'ies'}.")


def store():
    keyring = _keyring()
    print("Storing SIP intranet credentials in Windows Credential Manager.")
    print("Use the same login you type at:")
    print(f"  {config.LOGON_URL}")
    print("The password is not echoed and is never written to a file.\n")

    username = input("TP intranet username: ").strip()
    if not username:
        print("Aborted: username cannot be empty.")
        sys.exit(1)

    password = getpass.getpass("TP intranet password: ")
    if not password:
        print("Aborted: password cannot be empty.")
        sys.exit(1)
    if getpass.getpass("Confirm password    : ") != password:
        print("Aborted: the passwords did not match.")
        sys.exit(1)

    keyring.set_password(config.KEYRING_SERVICE_NAME, config.KEYRING_USERNAME_KEY, username)
    keyring.set_password(config.KEYRING_SERVICE_NAME, config.KEYRING_PASSWORD_KEY, password)

    stored_user, stored_password = config._credentials_from_keyring()
    if stored_user != username or stored_password != password:
        print("\nSaved, but reading it back gave a different value. Check Credential Manager.")
        sys.exit(1)

    print(f"\nStored. Credentials for '{username}' are in Windows Credential Manager")
    print(f"under the service name '{config.KEYRING_SERVICE_NAME}'.")
    print("\nVerify with a rehearsal that creates no record:")
    print("  python run_upload_file.py --dry-run --keep-open")
    _warn_if_env_shadows()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--show", action="store_true", help="Show what is configured (never the password)."
    )
    group.add_argument("--clear", action="store_true", help="Remove the stored credentials.")
    args = parser.parse_args()

    if args.show:
        show()
    elif args.clear:
        clear()
    else:
        store()


if __name__ == "__main__":
    main()
