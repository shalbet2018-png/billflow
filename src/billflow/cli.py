import argparse
import os
import sys

from .config import ConfigError, load_config
from .engine import commit, plan
from .models import BillingError


def _show(results):
    print(f"\n{'Client':<12}{'Name':<24}{'Status':<9}{'Invoice':<12}{'Total':>12}  Notes")
    for r in results:
        total = f"{r.invoice.total:,.2f}" if r.invoice else ""
        print(f"{r.client_id:<12}{r.client_name[:22]:<24}{r.status:<9}{(r.invoice.number if r.invoice else ''):<12}{total:>12}  {r.message}")
    ok = sum(r.status == "OK" for r in results)
    print(f"\n{ok} ready, {sum(r.status == 'ERROR' for r in results)} need attention, "
          f"{sum(r.status == 'SKIPPED' for r in results)} already done.")


def main(argv=None):
    ap = argparse.ArgumentParser(prog="billflow", description="Monthly invoicing and ledger automation")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p_init = sub.add_parser("init", help="create a demo workspace")
    p_init.add_argument("folder")
    p_run = sub.add_parser("run", help="prepare invoices for a month")
    p_run.add_argument("--config", required=True)
    p_run.add_argument("--period", required=True, help="billing month, e.g. 2026-11")
    p_run.add_argument("--dry-run", action="store_true", help="show what would happen, write nothing")
    p_run.add_argument("--yes", action="store_true", help="skip the confirmation question")
    p_run.add_argument("--send-emails", action="store_true",
                       help="actually send emails (otherwise they are only saved as .eml drafts)")
    p_run.add_argument("--push", action="store_true",
                       help="actually send invoices to Xero/QuickBooks/webhook outputs set to mode: api")
    p_run.add_argument("--force", action="store_true", help="re-issue clients already invoiced for the period")
    p_run.add_argument("--only", default="", help="comma-separated client ids")
    sub.add_parser("gui", help="open the desktop screen")
    p_set = sub.add_parser("setup", help="guided setup of a new business (opens the wizard)")
    p_set.add_argument("--answers", help="JSON file of answers: write the setup without opening the screen")
    p_set.add_argument("--folder", help="where to save the setup (with --answers)")
    p_lic = sub.add_parser("license", help="show or install a licence key")
    lsub = p_lic.add_subparsers(dest="lcmd", required=True)
    lsub.add_parser("status", help="show the active licence")
    p_inst = lsub.add_parser("install", help="install a licence key (the text or a file containing it)")
    p_inst.add_argument("key")
    args = ap.parse_args(argv)

    if args.cmd == "gui":
        from .gui import main as gui_main
        return gui_main()

    if args.cmd == "setup":
        if args.answers:
            import json
            from . import wizard
            if not args.folder:
                print("Error: --folder is required with --answers", file=sys.stderr)
                return 2
            try:
                with open(args.answers, encoding="utf-8") as f:
                    print("Saved " + str(wizard.write_config(json.load(f), args.folder)))
            except (OSError, ValueError, ConfigError) as e:
                print(f"Error: {e}", file=sys.stderr)
                return 2
            return 0
        import tkinter as tk
        from .wizard_gui import open_wizard
        from .gui import set_window_icon
        root = tk.Tk()
        set_window_icon(root)
        root.withdraw()
        open_wizard(root, lambda path, period: print(path)).win.wait_window()
        root.destroy()
        return 0

    if args.cmd == "license":
        from . import licensing
        if args.lcmd == "status":
            print(licensing.current().describe())
            return 0
        key = args.key
        if os.path.isfile(key):
            key = open(key, encoding="utf-8").read()
        try:
            print("Installed. " + licensing.install(key).describe())
        except licensing.LicenseError as e:
            print(f"Error: {e}", file=sys.stderr)
            return 2
        return 0

    if args.cmd == "init":
        from .sample import init_workspace
        t = init_workspace(args.folder)
        print(f"Demo workspace created in {t}\nTry: billflow run --config \"{t}\\config.yml\" --period 2026-11 --dry-run")
        return 0

    try:
        cfg = load_config(args.config)
        only = [x.strip() for x in args.only.split(",") if x.strip()] or None
        results = plan(cfg, args.period, only, args.force)
    except (ConfigError, BillingError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2
    _show(results)
    from . import licensing
    lic = licensing.current()
    n_ok = sum(r.status == "OK" for r in results)
    if not lic.valid or lic.max_clients:
        print(lic.describe())
    if lic.max_clients and n_ok > lic.max_clients and not args.dry_run:
        print(f"Error: this run has {n_ok} clients but your licence allows {lic.max_clients}. "
              "Nothing was written. Contact your supplier to upgrade.", file=sys.stderr)
        return 4
    if args.dry_run or not any(r.status == "OK" for r in results):
        print("Dry run - nothing was written." if args.dry_run else "Nothing to do.")
        return 0
    if not args.yes and input("\nCreate these invoices now? [y/N] ").strip().lower() != "y":
        print("Cancelled - nothing was written.")
        return 1
    ok = [r for r in results if r.status == "OK"]
    cfg["_send_emails"] = bool(args.send_emails)
    cfg["_push"] = bool(args.push)
    live = [o["type"] for o in cfg["outputs"] if o["type"] in ("xero", "quickbooks", "webhook") and o.get("mode") == "api"]
    if args.push and live:
        print(f"\nPUSH: {len(ok)} invoice(s) will be sent to: {', '.join(sorted(set(live)))}.")
        if not args.yes and input("Send to these systems for real? [y/N] ").strip().lower() != "y":
            cfg["_push"] = False
            print("Invoices will only be saved as JSON files.")
    if args.send_emails and any(o["type"] == "email" for o in cfg["outputs"]):
        print(f"\nEMAIL: {len(ok)} email(s) will be SENT to clients.")
        if not args.yes and input("Send emails for real? [y/N] ").strip().lower() != "y":
            cfg["_send_emails"] = False
            print("Emails will only be saved as drafts.")
    summary = commit(cfg, args.period, results)
    print(f"Done. Summary: {summary}")
    return 0 if all(r.status != "ERROR" for r in results) else 3


if __name__ == "__main__":
    sys.exit(main())
