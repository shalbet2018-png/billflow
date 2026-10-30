"""BILLFLOW desktop screen (tkinter): pick a config, preview, confirm, run. No command line needed."""
import json
import os
import queue
import random
import re
import threading
import tkinter as tk
import tkinter.font as tkfont
from datetime import date
from tkinter import filedialog, messagebox

from . import licensing
from .config import ConfigError, load_config, resolve_path
from .engine import commit, plan
from .models import BillingError

BG, PANEL = "#05060f", "#080a1c"
CYAN, MAG, GREEN, YEL, RED = "#00f0ff", "#ff2bd6", "#00ff9c", "#fcee09", "#ff3860"
FONT = "Consolas"
W, H = 1040, 720


def settings_path():
    return os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "billflow", "gui.json")


def load_settings():
    try:
        with open(settings_path(), encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {}


def save_settings(d):
    try:
        os.makedirs(os.path.dirname(settings_path()), exist_ok=True)
        with open(settings_path(), "w", encoding="utf-8") as fh:
            json.dump(d, fh)
    except OSError:
        pass


def format_results(results):
    """Plain-text table used by the log panel: list of (text, tag)."""
    rows = [(f"{'CLIENT':<11}{'NAME':<24}{'STATUS':<9}{'INVOICE':<12}{'TOTAL':>13}  NOTES", "head")]
    for r in results:
        total = f"{r.invoice.total:,.2f}" if r.invoice else ""
        rows.append((f"{r.client_id:<11}{r.client_name[:22]:<24}{r.status:<9}{(r.invoice.number if r.invoice else ''):<12}"
                     f"{total:>13}  {r.message}", r.status.lower()))
    return rows


def summarize(results):
    ok = [r for r in results if r.status == "OK"]
    return {"clients": len(results), "ready": len(ok), "errors": sum(r.status == "ERROR" for r in results),
            "total": sum((r.invoice.total for r in ok), 0)}


def ticker_text(results, currency=""):
    bits = [f"INV {r.invoice.number} :: {currency} {r.invoice.total:,.2f} :: DEBIT POSTED" for r in results if r.invoice]
    return "   //   ".join(bits) or "LEDGER IDLE :: AWAITING DATA :: ALL ACCOUNTS BALANCED :: NO ANOMALIES"


class Belt:
    """Rows of binary digits sliding left to right like a conveyor belt, behind everything."""

    def __init__(self, canvas, width, height):
        self.c, self.w, self.h = canvas, width, height
        self.font = tkfont.Font(family=FONT, size=11)
        self.lanes = []
        shades = ["#06222a", "#07303a", "#0a3d3a", "#0b2f4a", "#1a1038"]
        for i, y in enumerate(range(8, height, 20)):
            speed = random.choice([1, 1, 2, 2, 3, 4])
            text = self._bits()
            wd = self.font.measure(text)
            fill = random.choice(shades)
            ids = [canvas.create_text(x, y, text=text, anchor="w", font=self.font, fill=fill, tags="belt")
                   for x in range(-wd, width + wd, wd)]
            self.lanes.append({"ids": ids, "speed": speed, "wd": wd, "y": y})
        canvas.tag_lower("belt")
        self.tick_n = 0

    @staticmethod
    def _bits(n=64):
        return "".join(random.choice("01") for _ in range(n))

    def step(self):
        for lane in self.lanes:
            for i in lane["ids"]:
                self.c.move(i, lane["speed"], 0)
                x0 = self.c.bbox(i)[0]
                if x0 > self.w:
                    self.c.move(i, -lane["wd"] * len(lane["ids"]), 0)
        self.tick_n += 1
        if self.tick_n % 45 == 0:
            lane = random.choice(self.lanes)
            for i in lane["ids"]:
                t = list(self.c.itemcget(i, "text"))
                for _ in range(6):
                    t[random.randrange(len(t))] = random.choice("01")
                self.c.itemconfigure(i, text="".join(t))


class App:
    def __init__(self, root):
        self.root = root
        root.title("BILLFLOW // NEURAL LEDGER")
        root.configure(bg=BG)
        root.geometry(f"{W}x{H}")
        root.resizable(False, False)
        self.c = tk.Canvas(root, width=W, height=H, bg=BG, highlightthickness=0)
        self.c.pack()
        self.belt = Belt(self.c, W, H)
        self.q = queue.Queue()
        self.busy = False
        self.last = None            # (key, cfg, results) from the latest preview
        s = load_settings()
        self.cfg_var = tk.StringVar(value=s.get("config", ""))
        self.period_var = tk.StringVar(value=date.today().strftime("%Y-%m"))
        self._build()
        self._tick()
        self._poll()
        self.refresh_license()

    # ---- layout ------------------------------------------------------------------------------------------------
    def _panel(self, x1, y1, x2, y2, color=CYAN):
        self.c.create_rectangle(x1, y1, x2, y2, fill=PANEL, outline=color, width=1)
        for (cx, cy, dx, dy) in ((x1, y1, 1, 1), (x2, y1, -1, 1), (x1, y2, 1, -1), (x2, y2, -1, -1)):
            self.c.create_line(cx, cy, cx + 14 * dx, cy, fill=color, width=3)
            self.c.create_line(cx, cy, cx, cy + 14 * dy, fill=color, width=3)

    def _btn(self, x, y, w, text, cmd, color=CYAN):
        b = tk.Button(self.root, text=text, command=cmd, font=(FONT, 10, "bold"), fg=color, bg=PANEL, bd=0,
                      activebackground=color, activeforeground=BG, relief="flat", cursor="hand2",
                      highlightthickness=1, highlightbackground=color, highlightcolor=color)
        b.bind("<Enter>", lambda e: b.configure(bg=color, fg=BG))
        b.bind("<Leave>", lambda e: b.configure(bg=PANEL, fg=color))
        self.c.create_rectangle(x - 1, y - 1, x + w, y + 30, outline=color, fill=PANEL)
        self.c.create_window(x, y, window=b, width=w, height=30, anchor="nw")
        return b

    def _build(self):
        c = self.c
        c.create_text(42, 44, text="BILLFLOW", anchor="w", font=(FONT, 38, "bold"), fill=CYAN)
        c.create_text(39, 42, text="BILLFLOW", anchor="w", font=(FONT, 38, "bold"), fill=MAG)
        c.create_text(44, 84, text="// NEURAL LEDGER INTERFACE  v0.1  //  ACCOUNTS BOT ONLINE", anchor="w",
                      font=(FONT, 10), fill=GREEN)
        self.lic_id = c.create_text(W - 40, 40, text="", anchor="e", font=(FONT, 10, "bold"), fill=YEL)
        self.sys_id = c.create_text(W - 40, 62, text="SYS: STANDBY", anchor="e", font=(FONT, 10), fill=GREEN)

        self._panel(30, 110, W - 30, 218)
        c.create_text(48, 126, text="[ SOURCE LINK ]", anchor="w", font=(FONT, 9, "bold"), fill=MAG)
        ent = dict(font=(FONT, 11), bg="#02030a", fg=CYAN, insertbackground=CYAN, relief="flat", highlightthickness=1,
                   highlightbackground="#0b5560", highlightcolor=CYAN)
        c.create_text(48, 154, text="CONFIG", anchor="w", font=(FONT, 10, "bold"), fill=GREEN)
        c.create_window(120, 142, window=tk.Entry(self.root, textvariable=self.cfg_var, **ent), width=650, height=26, anchor="nw")
        self._btn(784, 140, 100, "BROWSE", self.browse)
        self._btn(894, 140, 100, "DEMO", self.make_demo, MAG)
        c.create_text(48, 192, text="PERIOD", anchor="w", font=(FONT, 10, "bold"), fill=GREEN)
        c.create_window(120, 180, window=tk.Entry(self.root, textvariable=self.period_var, **ent), width=110, height=26, anchor="nw")
        c.create_text(240, 192, text="YYYY-MM", anchor="w", font=(FONT, 8), fill="#3a6a70")
        self._btn(300, 178, 190, "SETUP WIZARD", self.open_wizard, MAG)
        self.b_prev = self._btn(500, 178, 150, "> PREVIEW", self.preview)
        self.b_run = self._btn(660, 178, 150, "> EXECUTE", self.execute, GREEN)
        self._btn(820, 178, 174, "OPEN OUTPUT", self.open_output, YEL)

        self.stat = {}
        for i, (key, label, color) in enumerate((("clients", "CLIENTS", CYAN), ("ready", "READY", GREEN),
                                                  ("errors", "FLAGGED", RED), ("total", "TOTAL BILLED", MAG))):
            x1 = 30 + i * 253
            self._panel(x1, 232, x1 + 238, 300, color)
            c.create_text(x1 + 16, 250, text=label, anchor="w", font=(FONT, 9, "bold"), fill=color)
            self.stat[key] = c.create_text(x1 + 16, 280, text="--", anchor="w", font=(FONT, 20, "bold"), fill="#ffffff")

        self._panel(30, 314, W - 30, 640, MAG)
        c.create_text(48, 330, text="[ LEDGER STREAM ]", anchor="w", font=(FONT, 9, "bold"), fill=CYAN)
        frame = tk.Frame(self.root, bg="#02030a", highlightthickness=1, highlightbackground="#3a0f38")
        self.log = tk.Text(frame, font=(FONT, 10), bg="#02030a", fg=GREEN, relief="flat", wrap="none", bd=0,
                           insertbackground=GREEN, state="disabled", padx=8, pady=6)
        sb = tk.Scrollbar(frame, command=self.log.yview, bg=PANEL, troughcolor=BG, bd=0)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(side="left", fill="both", expand=True)
        c.create_window(44, 344, window=frame, width=W - 88, height=284, anchor="nw")
        for tag, col in (("head", CYAN), ("ok", GREEN), ("error", RED), ("skipped", YEL), ("info", "#8aa7ff"), ("warn", YEL)):
            self.log.tag_configure(tag, foreground=col)
        self.log.tag_configure("head", underline=True)

        c.create_rectangle(0, H - 62, W, H, fill="#02030a", outline="", tags="ticker_bg")
        c.create_line(0, H - 62, W, H - 62, fill=CYAN)
        c.create_line(0, H - 2, W, H - 2, fill=MAG)
        self.tfont = tkfont.Font(family=FONT, size=12, weight="bold")
        self.set_ticker(ticker_text([]))
        self.say("BILLFLOW online. Pick a config (or press DEMO), choose the month and press PREVIEW. "
                 "Nothing is written until you press EXECUTE and confirm.", "info")

    # ---- belt + ticker animation -------------------------------------------------------------------------------
    def set_ticker(self, text):
        for i in getattr(self, "ticker_ids", []):
            self.c.delete(i)
        text = "  >>>  " + text + "  <<<  "
        wd = self.tfont.measure(text)
        self.ticker_wd = wd
        self.ticker_ids = [self.c.create_text(x, H - 32, text=text, anchor="w", font=self.tfont, fill=CYAN)
                           for x in range(-wd, W + wd, wd)]

    def _tick(self):
        self.belt.step()
        for i in self.ticker_ids:
            self.c.move(i, 3, 0)
            if self.c.bbox(i)[0] > W:
                self.c.move(i, -self.ticker_wd * len(self.ticker_ids), 0)
        self.c.tag_raise("ticker_bg")
        for i in self.ticker_ids:
            self.c.tag_raise(i)
        self.root.after(33, self._tick)

    # ---- helpers -----------------------------------------------------------------------------------------------
    def say(self, text, tag="info"):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def set_sys(self, text, color=GREEN):
        self.c.itemconfigure(self.sys_id, text="SYS: " + text, fill=color)

    def refresh_license(self):
        lic = licensing.current()
        if lic.valid:
            text, color = f"LICENSED // {lic.customer.upper()[:28]} // EXP {lic.expires:%Y-%m-%d}", GREEN
        else:
            text, color = f"TRIAL MODE // MAX {licensing.TRIAL_MAX_CLIENTS} CLIENTS PER RUN", YEL
        self.c.itemconfigure(self.lic_id, text=text, fill=color)

    def set_stats(self, results):
        s = summarize(results)
        self.c.itemconfigure(self.stat["clients"], text=str(s["clients"]))
        self.c.itemconfigure(self.stat["ready"], text=str(s["ready"]))
        self.c.itemconfigure(self.stat["errors"], text=str(s["errors"]))
        self.c.itemconfigure(self.stat["total"], text=f"{s['total']:,.2f}")

    def _busy(self, on):
        self.busy = on
        for b in (self.b_prev, self.b_run):
            b.configure(state="disabled" if on else "normal")
        self.set_sys("PROCESSING..." if on else "STANDBY", MAG if on else GREEN)

    def _key(self):
        return (self.cfg_var.get().strip(), self.period_var.get().strip())

    # ---- actions -----------------------------------------------------------------------------------------------
    def browse(self):
        p = filedialog.askopenfilename(title="Choose config.yml", filetypes=[("Config", "*.yml *.yaml"), ("All", "*.*")])
        if p:
            self.cfg_var.set(p)

    def open_wizard(self):
        from .wizard_gui import open_wizard

        def done(path, period):
            self.cfg_var.set(path)
            self.period_var.set(period)

        open_wizard(self.root, done)

    def make_demo(self):
        from .sample import init_workspace
        d = filedialog.askdirectory(title="Choose an empty folder for the demo workspace")
        if not d:
            return
        try:
            t = init_workspace(os.path.join(d, "billflow-demo"))
        except OSError as e:
            messagebox.showerror("BILLFLOW", f"Could not create the demo: {e}")
            return
        self.cfg_var.set(os.path.join(str(t), "config.yml"))
        self.period_var.set("2026-11")
        self.say(f"Demo workspace created in {t} (fake data). Press PREVIEW.", "info")

    def preview(self):
        if self.busy:
            return
        path, period = self._key()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("BILLFLOW", "Choose a config.yml first (BROWSE, or DEMO to try it with sample data).")
            return
        if not re.fullmatch(r"\d{4}-\d{2}", period):
            messagebox.showwarning("BILLFLOW", "Period must look like 2026-11.")
            return
        save_settings({"config": path})
        self.clear_log()
        self.say(f"> scanning sources for {period} ...", "info")
        self._busy(True)

        def work():
            try:
                cfg = load_config(path)
                self.q.put(("previewed", (path, period), cfg, plan(cfg, period)))
            except (ConfigError, BillingError, OSError) as e:
                self.q.put(("error", str(e)))
            except Exception as e:  # keep the screen alive on unexpected faults
                self.q.put(("error", f"Unexpected problem: {e}"))
        threading.Thread(target=work, daemon=True).start()

    def execute(self):
        if self.busy:
            return
        if not self.last or self.last[0] != self._key():
            messagebox.showinfo("BILLFLOW", "Press PREVIEW first, so you can see exactly what will be created.")
            return
        _, cfg, results = self.last
        ok = [r for r in results if r.status == "OK"]
        if not ok:
            messagebox.showinfo("BILLFLOW", "Nothing is ready to create.")
            return
        lic = licensing.current()
        if lic.max_clients and len(ok) > lic.max_clients:
            messagebox.showerror("BILLFLOW", f"This run has {len(ok)} clients but your licence allows {lic.max_clients}. "
                                             "Nothing was written.")
            return
        total = sum(r.invoice.total for r in ok)
        if not messagebox.askyesno("CONFIRM", f"Create {len(ok)} invoice(s), total {total:,.2f}?\n\n"
                                              "Files will be written and invoice numbers used up. "
                                              "Emails and online systems are NOT contacted from this screen."):
            self.say("> cancelled - nothing was written.", "warn")
            return
        cfg["_send_emails"] = False
        cfg["_push"] = False
        self._busy(True)
        self.say("> writing files ...", "info")
        period = self._key()[1]

        def work():
            try:
                self.q.put(("committed", commit(cfg, period, results), results))
            except (ConfigError, BillingError, OSError) as e:
                self.q.put(("error", str(e)))
            except Exception as e:
                self.q.put(("error", f"Unexpected problem: {e}"))
        threading.Thread(target=work, daemon=True).start()

    def open_output(self):
        if not self.last:
            path = self.cfg_var.get().strip()
            if not os.path.isfile(path):
                messagebox.showinfo("BILLFLOW", "Choose a config first.")
                return
            cfg = load_config(path)
        else:
            cfg = self.last[1]
        out = resolve_path(cfg, cfg.get("output_dir", "output"))
        if out.exists():
            os.startfile(str(out))
        else:
            messagebox.showinfo("BILLFLOW", "No output yet - run EXECUTE first.")

    def _poll(self):
        try:
            while True:
                m = self.q.get_nowait()
                if m[0] == "error":
                    self._busy(False)
                    self.set_sys("FAULT", RED)
                    self.say("ERROR: " + m[1], "error")
                elif m[0] == "previewed":
                    _, key, cfg, results = m
                    self.last = (key, cfg, results)
                    self._busy(False)
                    for text, tag in format_results(results):
                        self.say(text, tag)
                    s = summarize(results)
                    self.say(f"\n{s['ready']} ready, {s['errors']} need attention. Nothing written yet.", "info")
                    self.set_stats(results)
                    self.set_ticker(ticker_text(results, cfg["company"].get("currency", "")))
                elif m[0] == "committed":
                    _, summary, results = m
                    self._busy(False)
                    self.last = None
                    done = sum(r.status == "OK" for r in results)
                    self.say(f"\n> DONE. {done} invoice set(s) written. Summary: {summary}", "ok")
                    for r in results:
                        if r.status == "ERROR":
                            self.say(f"  ! {r.client_id}: {r.message}", "error")
        except queue.Empty:
            pass
        self.root.after(100, self._poll)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    main()

