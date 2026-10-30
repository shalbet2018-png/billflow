"""BILLFLOW desktop screen (tkinter). Made for people who are not technical:

1. choose your business setup   2. choose the month   3. Check invoices   4. Create invoices

Nothing is written until Create invoices is pressed and confirmed. No command line is needed for anything here,
including the licence key. Two looks: Neon (the original) and Plain; the choice is remembered.
"""
import calendar
import json
import os
import queue
import random
import subprocess
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox

from . import licensing
from .config import ConfigError, load_config, parse_period, resolve_path
from .engine import commit, plan
from .friendly import describe_results, explain, month_label
from .models import BillingError

# ---- looks -------------------------------------------------------------------------------------------------------
THEMES = {
    "neon": dict(BG="#05060f", PANEL="#080a1c", CYAN="#00f0ff", MAG="#ff2bd6", GREEN="#00ff9c", YEL="#fcee09",
                 RED="#ff3860", TEXT="#c9f6ff", MUTED="#5a8f96", INPUT="#02030a", INPUT_FG="#00f0ff", BORDER="#0b5560",
                 BTN_BG="#080a1c", HOVER_FG="#05060f", FONT="Consolas"),
    "plain": dict(BG="#f3f5f8", PANEL="#ffffff", CYAN="#1d5fbf", MAG="#5b4bb7", GREEN="#1b7a3a", YEL="#8a5a00",
                  RED="#c0392b", TEXT="#1f2933", MUTED="#5f6b7a", INPUT="#ffffff", INPUT_FG="#1f2933", BORDER="#c8d0da",
                  BTN_BG="#eef2f7", HOVER_FG="#ffffff", FONT="Segoe UI"),
}
THEME = "neon"
BG = PANEL = CYAN = MAG = GREEN = YEL = RED = TEXT = MUTED = INPUT = INPUT_FG = BORDER = BTN_BG = HOVER_FG = FONT = None


def apply_theme(name):
    """Switches every colour/font name in this module (and in the wizard screen, if loaded)."""
    global THEME
    THEME = name if name in THEMES else "neon"
    g = globals()
    g.update(THEMES[THEME])
    wiz = sys.modules.get("billflow.wizard_gui")
    if wiz is not None:
        wiz.sync_theme()


apply_theme("neon")
W, H = 1060, 760


# ---- settings ----------------------------------------------------------------------------------------------------
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


def open_path(path):
    """Opens a file or folder with whatever the computer normally uses for it."""
    path = str(path)
    try:
        if hasattr(os, "startfile"):
            os.startfile(path)                              # Windows
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except OSError as e:
        messagebox.showerror("BILLFLOW", f"Could not open {path}\n\n{e}")


# ---- small helpers -----------------------------------------------------------------------------------------------
def summarize(results):
    ok = [r for r in results if r.status == "OK"]
    return {"clients": len(results), "ready": len(ok), "errors": sum(r.status == "ERROR" for r in results),
            "skipped": sum(r.status == "SKIPPED" for r in results), "total": sum((r.invoice.total for r in ok), 0)}


def month_choices(today=None, back=24, ahead=3):
    """Months to offer, newest first: a few ahead of today and two years back."""
    today = today or date.today()
    out = []
    for k in range(ahead, -back - 1, -1):
        n = today.year * 12 + today.month - 1 + k
        out.append(f"{n // 12}-{n % 12 + 1:02d}")
    return out


def setup_label(path):
    """How a setup is shown in the list: the company name, and the folder it lives in."""
    try:
        import yaml
        with open(path, encoding="utf-8") as fh:
            name = ((yaml.safe_load(fh) or {}).get("company") or {}).get("name") or ""
    except Exception:  # unreadable setup: show its path instead
        name = ""
    folder = Path(path).parent.name
    return f"{name}   ({folder})" if name else str(path)


def plural(n, word):
    return f"{n} {word}" + ("" if n == 1 else "s")


def styled_button(parent, text, cmd, color=None, big=False, width=None):
    color = color or CYAN
    b = tk.Button(parent, text=text, command=cmd, font=(FONT, 12 if big else 10, "bold"), fg=color, bg=BTN_BG, bd=0,
                  activebackground=color, activeforeground=HOVER_FG, relief="flat", cursor="hand2", padx=14,
                  pady=7 if big else 4, highlightthickness=1, highlightbackground=color, highlightcolor=color,
                  disabledforeground=MUTED, width=width)

    def enter(_):
        if str(b["state"]) != "disabled":
            b.configure(bg=color, fg=HOVER_FG)

    b.bind("<Enter>", enter)
    b.bind("<Leave>", lambda e: b.configure(bg=BTN_BG, fg=color))
    return b


def style_menu(menu, width=30):
    menu.configure(font=(FONT, 11), bg=INPUT, fg=INPUT_FG, activebackground=CYAN, activeforeground=HOVER_FG,
                   relief="flat", highlightthickness=1, highlightbackground=BORDER, width=width, anchor="w", bd=0)
    menu["menu"].configure(bg=INPUT, fg=INPUT_FG, font=(FONT, 11), activebackground=CYAN, activeforeground=HOVER_FG)


class MonthMenu(tk.Frame):
    """Pick a month from a list ('November 2026'); var holds it as 2026-11 for the engine."""

    def __init__(self, master, var, width=18, bg=None):
        super().__init__(master, bg=bg or PANEL)
        self.var, self.shown = var, tk.StringVar()
        self.choices = month_choices()
        self.menu = tk.OptionMenu(self, self.shown, "")
        style_menu(self.menu, width)
        self.menu.pack()
        self._fill()
        self._trace = var.trace_add("write", lambda *a: self._sync())
        self.bind("<Destroy>", self._untrace)
        self._sync()

    def _untrace(self, e):
        if e.widget is self:
            try:
                self.var.trace_remove("write", self._trace)
            except tk.TclError:
                pass

    def _fill(self):
        m = self.menu["menu"]
        m.delete(0, "end")
        for p in self.choices:
            m.add_command(label=month_label(p), command=lambda v=p: self.var.set(v))

    def _sync(self):
        p = self.var.get().strip()
        try:
            parse_period(p)
        except ConfigError:
            self.shown.set("Choose a month")
            return
        if p not in self.choices:
            self.choices = sorted(set(self.choices + [p]), reverse=True)
            self._fill()
        self.shown.set(month_label(p))


class Belt:
    """Rows of binary digits sliding left to right like a conveyor belt (Neon look only, behind the title)."""

    def __init__(self, canvas, width, height):
        self.c, self.w, self.h = canvas, width, height
        self.font = tkfont.Font(family=FONT, size=11)
        self.lanes = []
        shades = ["#06222a", "#07303a", "#0a3d3a", "#0b2f4a", "#1a1038"]
        for y in range(8, height, 20):
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


# ---- licence key window ------------------------------------------------------------------------------------------
class LicenceDialog(tk.Toplevel):
    def __init__(self, master, on_done=None):
        super().__init__(master)
        self.on_done = on_done
        self.title("Enter licence key")
        self.configure(bg=PANEL)
        self.geometry("620x380")
        self.transient(master)
        lic = licensing.current()
        if lic.valid:
            now = lic.describe()
        else:
            now = f"You are on the free trial: up to {licensing.TRIAL_MAX_CLIENTS} invoices per run."
            if lic.note and not lic.note.startswith("To add"):
                now += " " + lic.note                                   # e.g. the installed key has expired
        tk.Label(self, text="Your licence", bg=PANEL, fg=CYAN, font=(FONT, 16, "bold"), anchor="w").pack(fill="x", padx=18, pady=(14, 0))
        tk.Label(self, text=now, bg=PANEL, fg=GREEN if lic.valid else YEL, font=(FONT, 10), anchor="w",
                 justify="left", wraplength=580).pack(fill="x", padx=18, pady=(2, 10))
        tk.Label(self, text="Paste the licence key we sent you (copy all of it), or open the key file:", bg=PANEL, fg=TEXT,
                 font=(FONT, 10), anchor="w").pack(fill="x", padx=18)
        self.box = tk.Text(self, height=6, font=(FONT, 10), bg=INPUT, fg=INPUT_FG, insertbackground=INPUT_FG, relief="flat",
                           highlightthickness=1, highlightbackground=BORDER, wrap="char")
        self.box.pack(fill="x", padx=18, pady=6)
        self.err = tk.Label(self, text="", bg=PANEL, fg=RED, font=(FONT, 10), anchor="w", justify="left", wraplength=580)
        self.err.pack(fill="x", padx=18)
        bar = tk.Frame(self, bg=PANEL)
        bar.pack(fill="x", padx=18, pady=12, side="bottom")
        styled_button(bar, "Save key", self.save, GREEN).pack(side="right")
        styled_button(bar, "Cancel", self.destroy, MUTED).pack(side="right", padx=8)
        styled_button(bar, "Open key file...", self.from_file).pack(side="left")
        self.grab_set()
        self.box.focus_set()

    def from_file(self):
        p = filedialog.askopenfilename(parent=self, title="Choose your licence key file",
                                       filetypes=[("Licence key", "*.key *.txt *.lic"), ("All files", "*.*")])
        if p:
            try:
                self.box.delete("1.0", "end")
                self.box.insert("1.0", Path(p).read_text(encoding="utf-8").strip())
            except (OSError, UnicodeDecodeError) as e:
                self.err.configure(text=f"Could not read that file: {e}")

    def save(self, quiet=False):
        key = "".join(self.box.get("1.0", "end").split())
        if not key:
            self.err.configure(text="Paste your licence key first.")
            return False
        try:
            lic = licensing.install(key)
        except licensing.LicenseError as e:
            text = str(e)
            if "not valid" in text:
                text = "That key was not accepted. Copy the whole key exactly as we sent it and try again."
            self.err.configure(text=text)
            return False
        if not quiet:
            messagebox.showinfo("BILLFLOW", "Licence saved.\n\n" + lic.describe(), parent=self)
        if self.on_done:
            self.on_done()
        self.destroy()
        return True


# ---- main screen -------------------------------------------------------------------------------------------------
HINTS = {
    "no_setup": "Start here: choose your business setup in step 1 (or press New setup).  "
                "Just looking? Press Try with sample data.",
    "ready_to_check": "Choose the month, then press Check invoices. Checking never saves or changes anything.",
    "busy": "Working... please wait.",
}


class App:
    def __init__(self, root):
        self.root = root
        root.title("BILLFLOW - monthly invoices")
        root.geometry(f"{W}x{H}")
        root.minsize(900, 640)
        self.settings = load_settings()
        apply_theme(self.settings.get("theme", "neon"))
        self.q = queue.Queue()
        self.busy = False
        self.last = None            # (key, cfg, results) from the latest check
        self.belt = None
        self.setups = [p for p in self.settings.get("setups", []) if os.path.isfile(p)]
        last_cfg = self.settings.get("config", "")
        if last_cfg and os.path.isfile(last_cfg) and last_cfg not in self.setups:
            self.setups.insert(0, last_cfg)
        self.cfg_var = tk.StringVar(value=last_cfg if os.path.isfile(last_cfg or "") else "")
        self.period_var = tk.StringVar(value=date.today().strftime("%Y-%m"))
        self.hint = tk.StringVar()
        self.cfg_var.trace_add("write", lambda *a: self._changed())
        self.period_var.trace_add("write", lambda *a: self._changed())
        self._build()
        self._tick()
        self._poll()

    # ---- layout --------------------------------------------------------------------------------------------------
    def _card(self, parent, step, title):
        outer = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=CYAN if THEME == "neon" else BORDER)
        head = tk.Frame(outer, bg=PANEL)
        head.pack(fill="x", padx=14, pady=(10, 4))
        if step:
            tk.Label(head, text=f" {step} ", bg=CYAN, fg=HOVER_FG, font=(FONT, 11, "bold")).pack(side="left")
        tk.Label(head, text=("  " if step else "") + title, bg=PANEL, fg=CYAN if THEME == "neon" else TEXT,
                 font=(FONT, 12, "bold")).pack(side="left")
        body = tk.Frame(outer, bg=PANEL)
        body.pack(fill="x", padx=14, pady=(0, 12))
        return outer, body

    def _build(self):
        root = self.root
        root.configure(bg=BG)
        for w in root.winfo_children():
            w.destroy()
        self.belt = None

        # header: name, licence, look switch
        head = tk.Frame(root, bg=BG)
        head.pack(fill="x", padx=24, pady=(14, 6))
        right = tk.Frame(head, bg=BG)
        right.pack(side="right", anchor="n")
        if THEME == "neon":
            self.c = tk.Canvas(head, width=600, height=86, bg=BG, highlightthickness=0)
            self.c.pack(side="left")
            self.belt = Belt(self.c, 600, 86)
            self.c.create_text(8, 34, text="BILLFLOW", anchor="w", font=(FONT, 34, "bold"), fill=CYAN)
            self.c.create_text(5, 32, text="BILLFLOW", anchor="w", font=(FONT, 34, "bold"), fill=MAG)
            self.c.create_text(10, 72, text="monthly invoices, checked before anything is saved", anchor="w",
                               font=(FONT, 10), fill=GREEN)
        else:
            self.c = None
            left = tk.Frame(head, bg=BG)
            left.pack(side="left")
            tk.Label(left, text="BILLFLOW", bg=BG, fg=CYAN, font=(FONT, 26, "bold")).pack(anchor="w")
            tk.Label(left, text="Monthly invoices, checked before anything is saved", bg=BG, fg=MUTED,
                     font=(FONT, 11)).pack(anchor="w")
        self.lic_label = tk.Label(right, text="", bg=BG, font=(FONT, 10, "bold"), anchor="e", justify="right")
        self.lic_label.pack(anchor="e")
        btns = tk.Frame(right, bg=BG)
        btns.pack(anchor="e", pady=(6, 0))
        styled_button(btns, "Enter licence key", self.enter_licence, YEL).pack(side="left")
        styled_button(btns, "Plain look" if THEME == "neon" else "Neon look", self.toggle_theme, MUTED).pack(side="left", padx=(8, 0))

        # step 1: setup
        top = tk.Frame(root, bg=BG)
        top.pack(fill="x", padx=24)
        card1, b1 = self._card(top, 1, "Your business setup")
        card1.pack(fill="x", pady=4)
        self.setup_shown = tk.StringVar()
        r1 = tk.Frame(b1, bg=PANEL)
        r1.pack(fill="x")
        self.setup_menu = tk.OptionMenu(r1, self.setup_shown, "")
        style_menu(self.setup_menu, 52)
        self.setup_menu.pack(side="left")
        styled_button(r1, "Find a setup...", self.browse).pack(side="left", padx=(10, 0))
        r2 = tk.Frame(b1, bg=PANEL)
        r2.pack(fill="x", pady=(8, 0))
        styled_button(r2, "New setup", self.open_wizard, MAG).pack(side="left")
        styled_button(r2, "Change my setup", self.change_setup, MAG).pack(side="left", padx=(8, 0))
        styled_button(r2, "Try with sample data", self.make_demo, MUTED).pack(side="left", padx=(8, 0))
        self._fill_setups()

        # steps 2-4
        row = tk.Frame(root, bg=BG)
        row.pack(fill="x", padx=24, pady=4)
        card2, b2 = self._card(row, 2, "Month to bill")
        card2.pack(side="left", fill="y")
        MonthMenu(b2, self.period_var, 18).pack(anchor="w")
        card3, b3 = self._card(row, "", "Then")
        card3.pack(side="left", fill="both", expand=True, padx=(10, 0))
        self.b_prev = styled_button(b3, "3   Check invoices", self.preview, CYAN, big=True)
        self.b_prev.pack(side="left")
        self.b_run = styled_button(b3, "4   Create invoices", self.execute, GREEN, big=True)
        self.b_run.pack(side="left", padx=10)
        self.b_out = styled_button(b3, "Open results folder", self.open_output, YEL, big=True)
        self.b_out.pack(side="left")

        # tiles
        tiles = tk.Frame(root, bg=BG)
        tiles.pack(fill="x", padx=24, pady=(6, 2))
        self.stat = {}
        for i, (key, label, color) in enumerate((("clients", "Clients", CYAN), ("ready", "Ready", GREEN),
                                                 ("errors", "Need attention", RED), ("total", "Total to bill", MAG))):
            t = tk.Frame(tiles, bg=PANEL, highlightthickness=1, highlightbackground=color if THEME == "neon" else BORDER)
            t.pack(side="left", fill="x", expand=True, padx=(0 if i == 0 else 8, 0))
            tk.Label(t, text=label, bg=PANEL, fg=color, font=(FONT, 10, "bold"), anchor="w").pack(fill="x", padx=12, pady=(8, 0))
            self.stat[key] = tk.StringVar(value="-")
            tk.Label(t, textvariable=self.stat[key], bg=PANEL, fg=TEXT, font=(FONT, 20, "bold"), anchor="w").pack(
                fill="x", padx=12, pady=(0, 8))

        # hint bar (bottom) and results
        bar = tk.Frame(root, bg=CYAN if THEME == "neon" else "#e6eefb")
        bar.pack(side="bottom", fill="x")
        tk.Label(bar, textvariable=self.hint, bg=bar["bg"], fg=HOVER_FG if THEME == "neon" else "#123a73",
                 font=(FONT, 11, "bold"), anchor="w", justify="left", wraplength=W - 60).pack(fill="x", padx=24, pady=8)
        res = tk.Frame(root, bg=PANEL, highlightthickness=1, highlightbackground=MAG if THEME == "neon" else BORDER)
        res.pack(fill="both", expand=True, padx=24, pady=(10, 12))
        tk.Label(res, text="Results", bg=PANEL, fg=CYAN if THEME == "neon" else TEXT, font=(FONT, 11, "bold"),
                 anchor="w").pack(fill="x", padx=12, pady=(8, 0))
        frame = tk.Frame(res, bg=INPUT)
        frame.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self.log = tk.Text(frame, font=(FONT, 11), bg=INPUT, fg=TEXT, relief="flat", wrap="word", bd=0,
                           state="disabled", padx=10, pady=8, spacing1=2, cursor="arrow")
        sb = tk.Scrollbar(frame, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(side="left", fill="both", expand=True)
        for tag, col in (("ok", GREEN), ("error", RED), ("skipped", YEL), ("info", TEXT), ("warn", YEL), ("muted", MUTED),
                         ("fix", TEXT)):
            self.log.tag_configure(tag, foreground=col)
        self.log.tag_configure("head", foreground=CYAN, font=(FONT, 12, "bold"), spacing1=10)
        self.log.tag_configure("name", font=(FONT, 11, "bold"))
        self.log.tag_configure("fix", lmargin1=28, lmargin2=28)
        self.log.tag_configure("indent", lmargin1=28, lmargin2=28)
        self._links = 0

        self.refresh_license()
        if self.last:
            self.show_results(*self.last[1:], key=self.last[0])
        else:
            self.say("Welcome. Choose your setup and the month, then press Check invoices. You will see every invoice "
                     "before anything is saved.", "info")
        self._refresh()

    # ---- state ---------------------------------------------------------------------------------------------------
    def _key(self):
        return (self.cfg_var.get().strip(), self.period_var.get().strip())

    def _checked(self):
        return bool(self.last) and self.last[0] == self._key()

    def _changed(self):
        p = self.cfg_var.get().strip()
        if not hasattr(self, "setup_menu"):
            return
        if p and os.path.isfile(p) and p not in self.setups:
            self.setups.insert(0, p)
            self._fill_setups()
        elif hasattr(self, "setup_shown"):
            self.setup_shown.set(setup_label(p) if p else "Choose your setup")
        self._refresh()

    def _refresh(self):
        """Enables what makes sense right now and says what to do next."""
        checked = self._checked()
        ready = checked and any(r.status == "OK" for r in self.last[2])
        if hasattr(self, "b_run"):
            self.b_prev.configure(state="disabled" if self.busy else "normal")
            self.b_run.configure(state="normal" if ready and not self.busy else "disabled")
        if self.busy:
            self.hint.set(HINTS["busy"])
        elif not self.cfg_var.get().strip():
            self.hint.set(HINTS["no_setup"])
        elif not checked:
            self.hint.set(HINTS["ready_to_check"])
        else:
            s = summarize(self.last[2])
            if s["errors"] and s["ready"]:
                verb = "needs" if s["errors"] == 1 else "need"
                self.hint.set(f"{plural(s['errors'], 'client')} {verb} attention (see What to do below). Fix them and press "
                              f"Check invoices again - or press Create invoices now for the {s['ready']} that are ready.")
            elif s["errors"]:
                self.hint.set("Nothing is ready yet - fix the problems listed below, then press Check invoices again.")
            elif s["ready"]:
                self.hint.set(f"All {s['ready']} look good. Press Create invoices to save them.")
            else:
                self.hint.set(f"Nothing new to create for {month_label(self._key()[1])} - every client has already been invoiced.")

    def _fill_setups(self):
        m = self.setup_menu["menu"]
        m.delete(0, "end")
        for p in self.setups:
            m.add_command(label=setup_label(p), command=lambda v=p: self.cfg_var.set(v))
        if not self.setups:
            m.add_command(label="(no setups yet - press New setup)", command=lambda: None)
        p = self.cfg_var.get().strip()
        self.setup_shown.set(setup_label(p) if p else "Choose your setup")

    def _remember(self, path):
        path = str(path)
        self.setups = [path] + [p for p in self.setups if p != path]
        self.settings.update(config=path, setups=self.setups[:12])
        save_settings(self.settings)

    def _busy(self, on):
        self.busy = on
        self._refresh()

    # ---- results panel -------------------------------------------------------------------------------------------
    def say(self, text, tag="info"):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _link(self, label, path):
        self._links += 1
        tag = f"link{self._links}"
        self.log.tag_configure(tag, foreground=CYAN, underline=True)
        self.log.tag_bind(tag, "<Button-1>", lambda e, p=path: open_path(p))
        self.log.tag_bind(tag, "<Enter>", lambda e: self.log.configure(cursor="hand2"))
        self.log.tag_bind(tag, "<Leave>", lambda e: self.log.configure(cursor="arrow"))
        self.log.configure(state="normal")
        self.log.insert("end", "      ", "indent")
        self.log.insert("end", label, tag)
        self.log.insert("end", "\n")
        self.log.configure(state="disabled")

    def show_problem(self, message, cfg=None, period=""):
        e = explain(message, cfg, period)
        self.say(e.text, "error")
        if e.fix:
            self.say("What to do: " + e.fix, "fix")
        if e.file:
            self._link(("Open folder " if os.path.isdir(e.file) else "Open ") + Path(e.file).name, e.file)

    def show_results(self, cfg, results, key=None):
        period = (key or self._key())[1]
        self.clear_log()
        cur = (cfg.get("company") or {}).get("currency", "")
        items = describe_results(results, cfg, period, cur)
        titles = {"ERROR": "Need attention", "SKIPPED": "Skipped", "OK": "Ready to create"}
        tags = {"ERROR": "error", "SKIPPED": "skipped", "OK": "ok"}
        self.say(f"{month_label(period)}  -  {(cfg.get('company') or {}).get('name', '')}", "head")
        for status in ("ERROR", "SKIPPED", "OK"):
            group = [d for d in items if d["status"] == status]
            if not group:
                continue
            self.say(f"{titles[status]} ({len(group)})", "head")
            for d in group:
                self.log.configure(state="normal")
                self.log.insert("end", "  • ", tags[status])
                self.log.insert("end", f"{d['client_id']}  {d['client_name']}", ("name", tags[status]))
                self.log.insert("end", f"   {d['text']}\n", tags[status] if status != "OK" else "info")
                self.log.configure(state="disabled")
                if d["fix"]:
                    self.say("What to do: " + d["fix"], "fix")
                if d["file"]:
                    self._link(("Open folder " if os.path.isdir(d["file"]) else "Open ") + Path(d["file"]).name, d["file"])
        self.say("")
        self.say("Nothing has been saved yet.", "muted")
        self.log.see("1.0")
        self.set_stats(results)

    def set_stats(self, results):
        s = summarize(results)
        self.stat["clients"].set(str(s["clients"]))
        self.stat["ready"].set(str(s["ready"]))
        self.stat["errors"].set(str(s["errors"]))
        self.stat["total"].set(f"{s['total']:,.2f}")

    def refresh_license(self):
        lic = licensing.current()
        if lic.valid:
            lim = "unlimited clients" if not lic.max_clients else f"up to {lic.max_clients} clients"
            text, color = f"Licensed to {lic.customer} ({lim})\nvalid until {lic.expires:%d %B %Y}", GREEN
        else:
            text, color = f"Trial: up to {licensing.TRIAL_MAX_CLIENTS} invoices per run\n(checking is always unlimited)", YEL
        self.lic_label.configure(text=text, fg=color)

    # ---- animation -----------------------------------------------------------------------------------------------
    def _tick(self):
        if self.belt is not None:
            try:
                self.belt.step()
            except tk.TclError:
                pass
        self.root.after(40, self._tick)

    # ---- actions -------------------------------------------------------------------------------------------------
    def toggle_theme(self):
        apply_theme("plain" if THEME == "neon" else "neon")
        self.settings["theme"] = THEME
        save_settings(self.settings)
        self._build()

    def enter_licence(self):
        return LicenceDialog(self.root, self.refresh_license)

    def browse(self):
        p = filedialog.askopenfilename(title="Find your setup file (config.yml)",
                                       filetypes=[("BILLFLOW setup", "*.yml *.yaml"), ("All files", "*.*")])
        if p:
            self.cfg_var.set(os.path.normpath(p))
            self._remember(self.cfg_var.get())

    def _wizard_done(self, path, period):
        self.last = None
        self.cfg_var.set(str(path))
        self._remember(path)
        self.period_var.set(period)
        self._fill_setups()
        self.say("Setup saved. Now press Check invoices.", "ok")

    def open_wizard(self):
        from .wizard_gui import open_wizard
        return open_wizard(self.root, self._wizard_done)

    def change_setup(self):
        path = self.cfg_var.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showinfo("BILLFLOW", "First choose the setup you want to change in step 1 "
                                            "(or press New setup to make one).")
            return None
        from .wizard_gui import open_wizard
        try:
            return open_wizard(self.root, self._wizard_done, config_path=path)
        except (ConfigError, OSError, ValueError) as e:
            messagebox.showerror("BILLFLOW", f"This setup could not be opened for changes:\n\n{e}")
            return None

    def make_demo(self):
        from .sample import init_workspace
        target = Path(os.path.expanduser("~")) / "Documents" / "BILLFLOW sample data"
        try:
            (target / "billflow-state.json").unlink(missing_ok=True)     # so the sample can be tried again
            t = init_workspace(target)
        except OSError as e:
            messagebox.showerror("BILLFLOW", f"Could not create the sample data: {e}")
            return
        self.last = None
        self.cfg_var.set(str(t / "config.yml"))
        self._remember(self.cfg_var.get())
        self.period_var.set("2026-11")
        self.clear_log()
        self.say(f"Sample data (made-up clients) is in {t}.", "info")
        self.say("Now press Check invoices to see what BILLFLOW would create.", "ok")

    def preview(self):
        if self.busy:
            return
        path, period = self._key()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("BILLFLOW", "First choose your business setup in step 1.\n\n"
                                               "No setup yet? Press New setup, or Try with sample data.")
            return
        try:
            parse_period(period)
        except ConfigError:
            messagebox.showwarning("BILLFLOW", "Choose the month to bill in step 2.")
            return
        self._remember(path)
        self.clear_log()
        self.say(f"Checking {month_label(period)} - reading your files...", "info")
        self._busy(True)

        def work():
            try:
                cfg = load_config(path)
                self.q.put(("previewed", (path, period), cfg, plan(cfg, period)))
            except (ConfigError, BillingError, OSError) as e:
                self.q.put(("error", str(e), period))
            except Exception as e:  # keep the screen alive on unexpected faults
                self.q.put(("error", f"Unexpected problem: {e}", period))
        threading.Thread(target=work, daemon=True).start()

    def execute(self):
        if self.busy:
            return
        if not self._checked():
            messagebox.showinfo("BILLFLOW", "Press Check invoices first, so you can see exactly what will be created.")
            return
        key, cfg, results = self.last
        ok = [r for r in results if r.status == "OK"]
        if not ok:
            messagebox.showinfo("BILLFLOW", "Nothing is ready to create.")
            return
        lic = licensing.current()
        if lic.max_clients and len(ok) > lic.max_clients:
            what = "Your trial" if not lic.valid else "Your licence"
            messagebox.showerror("BILLFLOW", f"{what} covers up to {lic.max_clients} invoices per run, and "
                                             f"{len(ok)} are ready.\n\nPress Enter licence key to add or upgrade your "
                                             "licence. Nothing was saved.")
            return
        total = sum(r.invoice.total for r in ok)
        cur = (cfg.get("company") or {}).get("currency", "")
        if not messagebox.askyesno("Create invoices?", f"Create {plural(len(ok), 'invoice')} for {month_label(key[1])}, "
                                                       f"total {cur} {total:,.2f}?\n\n"
                                                       "They are saved in your results folder and the invoice numbers are "
                                                       "used up. Nothing is emailed or sent online from this screen."):
            self.say("Cancelled - nothing was saved.", "warn")
            return
        cfg["_send_emails"] = False
        cfg["_push"] = False
        self._busy(True)
        self.say("Saving invoices...", "info")

        def work():
            try:
                self.q.put(("committed", commit(cfg, key[1], results), cfg, results, key))
            except (ConfigError, BillingError, OSError) as e:
                self.q.put(("error", str(e), key[1]))
            except Exception as e:
                self.q.put(("error", f"Unexpected problem: {e}", key[1]))
        threading.Thread(target=work, daemon=True).start()

    def _results_folder(self, cfg=None, period=None):
        if cfg is None:
            path = self.cfg_var.get().strip()
            if not os.path.isfile(path):
                return None
            cfg = load_config(path)
        out = resolve_path(cfg, cfg.get("output_dir", "output"))
        sub = out / (period or self._key()[1])
        return sub if sub.exists() else out

    def open_output(self):
        try:
            out = self._results_folder(self.last[1] if self.last else None)
        except ConfigError:
            out = None
        if out is None:
            messagebox.showinfo("BILLFLOW", "First choose your business setup in step 1.")
        elif out.exists():
            open_path(out)
        else:
            messagebox.showinfo("BILLFLOW", "There are no results yet - press Check invoices, then Create invoices.")

    def _poll(self):
        try:
            while True:
                m = self.q.get_nowait()
                if m[0] == "error":
                    self._busy(False)
                    self.clear_log()
                    self.say("Something needs fixing before invoices can be checked:", "head")
                    cfg = None
                    try:
                        cfg = load_config(self._key()[0])
                    except (ConfigError, OSError, ValueError):
                        pass
                    self.show_problem(m[1], cfg, m[2])
                    self.hint.set("Something needs fixing - see What to do in the results below.")
                elif m[0] == "previewed":
                    _, key, cfg, results = m
                    self.last = (key, cfg, results)
                    self._busy(False)
                    self.show_results(cfg, results, key)
                    self._refresh()
                elif m[0] == "committed":
                    _, summary, cfg, results, key = m
                    self.last = None
                    self._busy(False)
                    done = sum(r.status == "OK" for r in results)
                    failed = [r for r in results if r.status == "ERROR" and r.message.startswith("Writing files failed")]
                    self.clear_log()
                    self.say(f"Done - {plural(done, 'invoice')} saved for {month_label(key[1])}.", "head")
                    folder = Path(summary).parent
                    self._link(f"Open the {month_label(key[1])} results folder", folder)
                    for r in failed:
                        e = explain(r.message, cfg, key[1], r.client_id)
                        self.say(f"  • {r.client_id}  {r.client_name}: {e.text}", "error")
                        self.say("What to do: " + e.fix, "fix")
                    self.hint.set(f"Done! {plural(done, 'invoice')} saved. Press Open results folder to see them.")
                    if done and messagebox.askyesno("BILLFLOW", f"Done - {plural(done, 'invoice')} saved for "
                                                                f"{month_label(key[1])}.\n\nOpen the results folder now?"):
                        open_path(folder)
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
