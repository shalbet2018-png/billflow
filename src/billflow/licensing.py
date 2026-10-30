"""Offline licence keys signed with Ed25519.

A key is  <payload>.<signature>  (both base64url). The vendor signs with a private key that is never shipped;
the matching public key below is embedded and can only verify, not create, keys.
Without a valid key billflow runs as a limited trial.
"""
import base64
import json
import os
from dataclasses import dataclass
from datetime import date
from pathlib import Path

PUBLIC_KEY_HEX = "dcd7d24be1d28eac943140431282a8afdc58de0ef8f652259e5119392985dc30"
TRIAL_MAX_CLIENTS = 5


class LicenseError(Exception):
    pass


@dataclass
class License:
    customer: str = "Trial"
    plan: str = "trial"
    expires: date = None
    max_clients: int = TRIAL_MAX_CLIENTS      # 0 = unlimited
    valid: bool = False
    note: str = ""

    def describe(self):
        if not self.valid:
            return f"Trial mode (max {TRIAL_MAX_CLIENTS} clients per run). {self.note}".strip()
        lim = "unlimited clients" if self.max_clients == 0 else f"up to {self.max_clients} clients"
        return f"Licensed to {self.customer} - {self.plan}, {lim}, expires {self.expires:%d %B %Y}"


def _b64d(s):
    return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))


def verify(key, today=None):
    from cryptography.exceptions import InvalidSignature
    from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey

    try:
        payload_b64, sig_b64 = key.strip().split(".")
        payload = _b64d(payload_b64)
        Ed25519PublicKey.from_public_bytes(bytes.fromhex(PUBLIC_KEY_HEX)).verify(_b64d(sig_b64), payload)
        d = json.loads(payload)
        expires = date.fromisoformat(d["expires"])
        lic = License(d["customer"], d.get("plan", "standard"), expires, int(d.get("max_clients", 0)), True)
    except (ValueError, KeyError, InvalidSignature) as e:
        raise LicenseError("Licence key is not valid.") from e
    if expires < (today or date.today()):
        raise LicenseError(f"Licence expired on {expires:%d %B %Y}.")
    return lic


def sign(private_pem, customer, plan, expires, max_clients):
    """Vendor-side only (see tools/make_license.py)."""
    from cryptography.hazmat.primitives import serialization

    priv = serialization.load_pem_private_key(private_pem, password=None)
    payload = json.dumps({"customer": customer, "plan": plan, "expires": expires,
                          "max_clients": max_clients}, separators=(",", ":")).encode()
    b = lambda x: base64.urlsafe_b64encode(x).decode().rstrip("=")
    return f"{b(payload)}.{b(priv.sign(payload))}"


def key_path():
    base = os.environ.get("APPDATA") or str(Path.home() / ".config")
    return Path(base) / "billflow" / "license.key"


def current(today=None):
    """The active licence: BILLFLOW_LICENSE env var, else the installed key file, else trial."""
    key = os.environ.get("BILLFLOW_LICENSE", "").strip()
    if not key and key_path().exists():
        key = key_path().read_text(encoding="utf-8").strip()
    if not key:
        return License(note="Install a key with: billflow license install <key>")
    try:
        return verify(key, today)
    except LicenseError as e:
        return License(note=str(e))


def install(key):
    lic = verify(key)
    p = key_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(key.strip(), encoding="utf-8")
    return lic
