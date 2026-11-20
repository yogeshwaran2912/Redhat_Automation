# Created by YOGESHWARAN TL
"""
Password lock for the Python source in this folder.

  lock.bat    (python secure_vault.py lock)
      Encrypt every .py file into vault\, compile it to a .pyc so the
      programs still run without a password, then delete the .py.
      The first time, it asks you to choose the password.

  unlock.bat  (python secure_vault.py unlock)
      Ask the password and put the .py files back so you can view / edit
      them. Run lock.bat again when you are done.

  python secure_vault.py status
      Show whether the folder is locked and how many tries are left.

3 wrong passwords -> no more tries for 24 hours.

Only the standard library is used: PBKDF2-SHA256 makes the key from the
password, HMAC-SHA256 in counter mode encrypts, and an HMAC tag catches a
wrong password or a damaged file.
"""
import base64
import datetime as dt
import getpass
import hashlib
import hmac
import json
import os
import py_compile
import secrets
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
VAULT = HERE / "vault"
VAULT_INFO = VAULT / "vault.json"
SELF = Path(__file__).name

MAX_TRIES = 3
LOCKOUT = dt.timedelta(hours=24)
PBKDF2_ROUNDS = 600_000

# The try counter is kept in two places, so deleting one doesn't reset it.
STATE_DIR = Path(os.getenv("LOCALAPPDATA", str(Path.home()))) / "SecureVault"


# ================================================================ crypto

def derive_keys(password, salt):
    key = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ROUNDS, dklen=64)
    return key[:32], key[32:]  # encryption key, MAC key


def keystream(enc_key, nonce, length):
    out = bytearray()
    counter = 0
    while len(out) < length:
        out += hmac.new(enc_key, nonce + counter.to_bytes(8, "big"), hashlib.sha256).digest()
        counter += 1
    return bytes(out[:length])


def encrypt(keys, data):
    enc_key, mac_key = keys
    nonce = secrets.token_bytes(16)
    body = bytes(a ^ b for a, b in zip(data, keystream(enc_key, nonce, len(data))))
    tag = hmac.new(mac_key, nonce + body, hashlib.sha256).digest()
    return nonce + tag + body


def decrypt(keys, blob):
    """-> bytes, or None when the password is wrong / the file is damaged."""
    enc_key, mac_key = keys
    nonce, tag, body = blob[:16], blob[16:48], blob[48:]
    if not hmac.compare_digest(tag, hmac.new(mac_key, nonce + body, hashlib.sha256).digest()):
        return None
    return bytes(a ^ b for a, b in zip(body, keystream(enc_key, nonce, len(body))))


# ================================================================ vault info + try counter

def load_info():
    if VAULT_INFO.exists():
        return json.loads(VAULT_INFO.read_text(encoding="utf-8"))
    return None


def save_info(info):
    VAULT.mkdir(exist_ok=True)
    VAULT_INFO.write_text(json.dumps(info, indent=2), encoding="utf-8")


def state_copy_path(info):
    return STATE_DIR / (info["salt"][:16].replace("/", "_").replace("+", "-") + ".json")


def load_tries(info):
    """Merge both copies of the counter, keeping the stricter one."""
    tries = {"fails": info.get("fails", 0), "locked_until": info.get("locked_until")}
    try:
        other = json.loads(state_copy_path(info).read_text(encoding="utf-8"))
        tries["fails"] = max(tries["fails"], other.get("fails", 0))
        if other.get("locked_until") and (not tries["locked_until"]
                                          or other["locked_until"] > tries["locked_until"]):
            tries["locked_until"] = other["locked_until"]
    except (OSError, ValueError):
        pass
    return tries


def save_tries(info, tries):
    info.update(tries)
    save_info(info)
    try:
        STATE_DIR.mkdir(parents=True, exist_ok=True)
        state_copy_path(info).write_text(json.dumps(tries), encoding="utf-8")
    except OSError:
        pass


def locked_out(tries):
    if tries["locked_until"]:
        until = dt.datetime.fromisoformat(tries["locked_until"])
        if dt.datetime.now() < until:
            return until
    return None


def ask_password(info):
    """Ask until right or out of tries -> keys, or None."""
    tries = load_tries(info)
    until = locked_out(tries)
    if until:
        print(f"Too many wrong passwords. Try again after {until:%d-%m-%Y %I:%M %p}.")
        return None
    if tries["locked_until"]:  # the 24 hours are over
        tries = {"fails": 0, "locked_until": None}
        save_tries(info, tries)

    salt = base64.b64decode(info["salt"])
    check = base64.b64decode(info["check"])
    while True:
        keys = derive_keys(getpass.getpass("Password: "), salt)
        if decrypt(keys, check) == b"ok":
            save_tries(info, {"fails": 0, "locked_until": None})
            return keys
        tries["fails"] += 1
        left = MAX_TRIES - tries["fails"]
        if left <= 0:
            until = dt.datetime.now() + LOCKOUT
            save_tries(info, {"fails": 0, "locked_until": until.isoformat(timespec="seconds")})
            print(f"Wrong password 3 times. Locked until {until:%d-%m-%Y %I:%M %p}.")
            return None
        save_tries(info, tries)
        print(f"Wrong password. {left} {'try' if left == 1 else 'tries'} left.")


def new_password():
    print("No password set yet - choose one (you need it to view or edit the code).")
    while True:
        pw = getpass.getpass("New password: ")
        if len(pw) < 4:
            print("Use at least 4 characters.")
            continue
        if getpass.getpass("Type it again: ") != pw:
            print("They don't match, try again.")
            continue
        salt = secrets.token_bytes(16)
        keys = derive_keys(pw, salt)
        info = {"salt": base64.b64encode(salt).decode(),
                "check": base64.b64encode(encrypt(keys, b"ok")).decode(),
                "fails": 0, "locked_until": None}
        save_info(info)
        return info, keys


# ================================================================ lock / unlock

def sources():
    return sorted(p for p in HERE.glob("*.py") if p.name != SELF)


def lock():
    files = sources()
    if not files:
        print("Nothing to lock - no .py files here (already locked?).")
        return 0

    info = load_info()
    if info:
        keys = ask_password(info)
        if not keys:
            return 1
    else:
        info, keys = new_password()

    # Compile everything first: a syntax error stops here with nothing changed.
    for src in files:
        try:
            py_compile.compile(str(src), cfile=str(src.with_suffix(".pyc")), doraise=True)
        except py_compile.PyCompileError as e:
            print(f"Can't compile {src.name} - fix it first, nothing was locked:\n{e.msg}")
            return 1

    for src in files:
        data = src.read_bytes()
        blob = encrypt(keys, data)
        if decrypt(keys, blob) != data:
            print(f"Encryption check failed for {src.name} - stopped, the .py was kept.")
            return 1
        (VAULT / (src.name + ".enc")).write_bytes(blob)
        src.unlink()
        print(f"  locked  {src.name}")

    shutil.rmtree(HERE / "__pycache__", ignore_errors=True)
    print(f"\nLocked {len(files)} files. The programs still run; use unlock.bat to view or edit.")
    return 0


def unlock():
    info = load_info()
    if not info:
        print("This folder has never been locked.")
        return 1
    blobs = sorted(VAULT.glob("*.py.enc"))
    if not blobs:
        print("The vault is empty.")
        return 1
    keys = ask_password(info)
    if not keys:
        return 1

    for blob in blobs:
        dst = HERE / blob.name[:-len(".enc")]
        data = decrypt(keys, blob.read_bytes())
        if data is None:
            print(f"  damaged {blob.name} - skipped")
            continue
        if dst.exists() and dst.read_bytes() != data:
            print(f"  kept    {dst.name} (already here and different - not overwritten)")
            continue
        dst.write_bytes(data)
        print(f"  opened  {dst.name}")

    print("\nThe .py files are back. Run lock.bat again when you finish editing.")
    return 0


def status():
    info = load_info()
    if not info:
        print("Not locked yet - run lock.bat to set a password.")
        return 0
    open_files = [p.name for p in sources()]
    print("Unlocked (.py files are visible): " + ", ".join(open_files) if open_files
          else "Locked - only .pyc files are here.")
    tries = load_tries(info)
    until = locked_out(tries)
    if until:
        print(f"Password blocked until {until:%d-%m-%Y %I:%M %p}.")
    else:
        print(f"Password tries left: {MAX_TRIES - tries['fails']}")
    return 0


if __name__ == "__main__":
    command = sys.argv[1] if len(sys.argv) > 1 else ""
    actions = {"lock": lock, "unlock": unlock, "status": status}
    if command not in actions:
        print(__doc__)
        sys.exit(2)
    try:
        sys.exit(actions[command]())
    except KeyboardInterrupt:
        print("\nCancelled.")
        sys.exit(1)
