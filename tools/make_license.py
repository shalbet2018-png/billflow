"""Vendor-only: create a customer licence key.

  python tools/make_license.py "Acme Properties" --expires 2027-10-31 --plan standard --max-clients 50

Needs the private signing key (default C:\\Users\\USER\\billflow-private\\signing-key.pem). NEVER commit that file.
"""
import argparse
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from billflow.licensing import sign, verify  # noqa: E402

ap = argparse.ArgumentParser()
ap.add_argument("customer")
ap.add_argument("--expires", required=True, help="YYYY-MM-DD")
ap.add_argument("--plan", default="standard")
ap.add_argument("--max-clients", type=int, default=0, help="0 = unlimited")
ap.add_argument("--private-key", default=os.environ.get("BILLFLOW_SIGNING_KEY", r"C:\Users\USER\billflow-private\signing-key.pem"))
a = ap.parse_args()
key = sign(Path(a.private_key).read_bytes(), a.customer, a.plan, a.expires, a.max_clients)
print(verify(key).describe(), file=sys.stderr)
print(key)
