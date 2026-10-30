"""BILLFLOW setup wizard screen: a business describes what it needs, links its own files and folders,
and the wizard writes the config. Nothing here touches the linked files except to read their column names."""
import os
import re
import tkinter as tk
from datetime import date
from tkinter import filedialog, messagebox

from . import gui, wizard
from .config import ConfigError
from .friendly import describe_results, month_label
from .models import BillingError

NONE = "(none)"
_THEME_NAMES = ("BG", "PANEL", "CYAN", "MAG", "GREEN", "YEL", "RED", "TEXT", "MUTED", "INPUT", "INPUT_FG", "BORDER",
                "BTN_BG", "HOVER_FG", "FONT")
BG = PANEL = CYAN = MAG = GREEN = YEL = RED = TEXT = MUTED = INPUT = INPUT_FG = BORDER = BTN_BG = HOVER_FG = FONT = None


def sync_theme():
    """Takes the current look (Neon or Plain) from the main screen."""
    globals().update({k: getattr(gui, k) for k in _THEME_NAMES})


sync_theme()


def _label(parent, text, color=None, size=10, bold=False, **kw):
    return tk.Label(parent, text=text, bg=PANEL, fg=color or TEXT, font=(FONT, size, "bold" if bold else "normal"), anchor="w",
                    justify="left", **kw)


def _entry(parent, var, width=30):
    return tk.Entry(parent, textvariable=var, width=width, font=(FONT, 11), bg=INPUT, fg=INPUT_FG, insertbackground=INPUT_FG,
                    relief="flat", highlightthickness=1, highlightbackground=BORDER, highlightcolor=CYAN)


def _number_entry(parent, var, width=12):
    """An entry that only accepts what can be part of a number (digits, one dot, commas)."""
    e = _entry(parent, var, width)
    ok = parent.register(lambda v: re.fullmatch(r"[0-9,]*\.?[0-9]*", v) is not None)
    e.configure(validate="key", validatecommand=(ok, "%P"))
    return e


def _num(text):
    """'1,500.50' -> Decimal, or None if it is not a plain number."""
    from decimal import Decimal, InvalidOperation
    t = str(text).replace(",", "").strip()
    if not re.fullmatch(r"\d+(\.\d+)?|\.\d+", t):
        return None
    try:
        return Decimal(t)
    except InvalidOperation:
        return None


def _button(parent, text, cmd, color=None, width=None):
    color = color or CYAN
    b = tk.Button(parent, text=text, command=cmd, font=(FONT, 10, "bold"), fg=color, bg=BTN_BG, bd=0, relief="flat",
                  activebackground=color, activeforeground=HOVER_FG, cursor="hand2", highlightthickness=1,
                  highlightbackground=color, highlightcolor=color, padx=10, pady=3, width=width)
    b.bind("<Enter>", lambda e: b.configure(bg=color, fg=HOVER_FG))
    b.bind("<Leave>", lambda e: b.configure(bg=BTN_BG, fg=color))
    return b


def _check(parent, text, var):
    return tk.Checkbutton(parent, text=text, variable=var, bg=PANEL, fg=TEXT, selectcolor=INPUT, activebackground=PANEL,
                          activeforeground=TEXT, font=(FONT, 10), anchor="w", highlightthickness=0, bd=0)


def _radio(parent, text, var, value, cmd=None):
    return tk.Radiobutton(parent, text=text, variable=var, value=value, command=cmd, bg=PANEL, fg=TEXT, selectcolor=INPUT,
                          activebackground=PANEL, activeforeground=TEXT, font=(FONT, 10), anchor="w", highlightthickness=0, bd=0)


class ColMenu(tk.Frame):
    """A drop-down of column names found in the linked file."""

    def __init__(self, master, allow_none=False, width=24):
        super().__init__(master, bg=PANEL)
        self.var = tk.StringVar()
        self.allow_none = allow_none
        self.menu = tk.OptionMenu(self, self.var, "")
        gui.style_menu(self.menu, width)
        self.menu.pack()
        self.set_options([])

    def set_options(self, options, guess=None):
        opts = ([NONE] if self.allow_none else []) + list(options)
        m = self.menu["menu"]
        m.delete(0, "end")
        for o in opts:
            m.add_command(label=o, command=lambda v=o: self.var.set(v))
        cur = self.var.get()
        if cur not in opts or (cur == NONE and guess in opts):
            self.var.set(guess if guess in opts else (opts[0] if opts else ""))

    def value(self):
        v = self.var.get()
        return "" if v == NONE else v

    def select(self, value):
        self.var.set(value or (NONE if self.allow_none else ""))


class FilePicker(tk.Frame):
    """Link a file: choose it, and the wizard reads its columns. Month in the name becomes a pattern."""

    def __init__(self, master, get_period, on_loaded=None, with_sheet=True):
        super().__init__(master, bg=PANEL)
        self.get_period, self.on_loaded = get_period, on_loaded
        self.file, self.sheet, self.hrow = tk.StringVar(), tk.StringVar(), tk.StringVar(value="1")
        self.status = tk.StringVar()
        top = tk.Frame(self, bg=PANEL)
        top.pack(fill="x")
        _entry(top, self.file, 58).pack(side="left", padx=(0, 6))
        _button(top, "Choose file...", self.browse).pack(side="left")
        row = tk.Frame(self, bg=PANEL)
        row.pack(fill="x", pady=(4, 0))
        if with_sheet:
            _label(row, "Sheet (Excel only)").pack(side="left")
            _entry(row, self.sheet, 18).pack(side="left", padx=(6, 12))
            _label(row, "Column names are in row").pack(side="left")
            e = _entry(row, self.hrow, 3)
            ok = row.register(lambda v: v.isdigit() or v == "")
            e.configure(validate="key", validatecommand=(ok, "%P"))
            e.pack(side="left", padx=(6, 12))
        _button(row, "Read columns", self.load, GREEN).pack(side="left")
        tk.Label(self, textvariable=self.status, bg=PANEL, fg=YEL, font=(FONT, 9), anchor="w", justify="left",
                 wraplength=700).pack(fill="x", pady=(2, 0))
        self.headers = []

    def browse(self):
        p = filedialog.askopenfilename(title="Choose the file", filetypes=[("Data files", "*.csv *.xlsx *.xlsm"), ("All", "*.*")])
        if not p:
            return
        period = self.get_period()
        self.file.set(wizard.suggest_template(os.path.normpath(p), period))
        self.sheet.set("")
        self.load()

    def load(self):
        period = self.get_period()
        try:
            info = wizard.inspect_file(wizard.fill(self.file.get(), period),
                                       wizard.fill(self.sheet.get(), period) if self.sheet.get().strip() else None,
                                       self.hrow.get() or 1)
        except (BillingError, ConfigError, OSError, ValueError) as e:
            from .friendly import explain
            self.status.set("Problem: " + explain(str(e), period=period).text)
            return False
        if info["sheets"] and not self.sheet.get().strip():
            self.sheet.set(wizard.suggest_template(info["sheets"][0], period))
        self.headers = info["headers"]
        pat = "  (the month in the name will follow each new month)" if self.file.get() != wizard.fill(self.file.get(), period) \
            or "{" in self.file.get() else ""
        self.status.set(f"OK - found {len(self.headers)} columns: {', '.join(self.headers[:10])}"
                        f"{'...' if len(self.headers) > 10 else ''}{pat}")
        if self.on_loaded:
            self.on_loaded(self.headers)
        return True

    def values(self):
        d = {"file": self.file.get().strip()}
        if self.sheet.get().strip():
            d["sheet"] = self.sheet.get().strip()
        if str(self.hrow.get()).strip() not in ("", "1"):
            d["header_row"] = int(self.hrow.get())
        return d

    def set_values(self, d):
        self.file.set(d.get("file", ""))
        self.sheet.set(d.get("sheet", ""))
        self.hrow.set(str(d.get("header_row", 1)))



class ChargeDialog(tk.Toplevel):
    """Add or change one charge. Asks only what that kind of charge needs."""

    KINDS = (("A fixed amount (rent, service charge)", "fixed"),
             ("Usage from meter readings (water, power)", "metered"),
             ("Extra bills added to a client (e.g. electricity bills)", "extra"),
             ("A percentage of other charges (VAT, late fee)", "percentage"))

    def __init__(self, wiz, existing=None, earlier=None):
        super().__init__(wiz.win)
        self.wiz, self.result = wiz, None
        self.earlier = list(earlier if earlier is not None else [c["name"] for c in wiz.charges])
        self.taken = {c["name"].strip().lower() for c in wiz.charges if c is not existing}
        self.title("Change a charge" if existing else "Add a charge")
        self.configure(bg=PANEL)
        self.geometry("820x600")
        self.transient(wiz.win)
        self.kind = tk.StringVar(value=(existing or {}).get("kind", "fixed"))
        self.name = tk.StringVar(value=(existing or {}).get("name", ""))
        tk.Frame(self, bg=PANEL, height=6).pack()
        _label(self, "  What kind of charge is it?", CYAN, 11, True).pack(fill="x")
        kinds = tk.Frame(self, bg=PANEL)
        kinds.pack(fill="x", padx=14)
        for text, val in self.KINDS:
            _radio(kinds, text, self.kind, val, self.rebuild).pack(fill="x")
        nm = tk.Frame(self, bg=PANEL)
        nm.pack(fill="x", padx=14, pady=8)
        _label(nm, "Name on the invoice  ").pack(side="left")
        _entry(nm, self.name, 36).pack(side="left")
        _label(nm, "  e.g. Rent, Water, VAT", MUTED, 9).pack(side="left")
        self.form = tk.Frame(self, bg=PANEL)
        self.form.pack(fill="both", expand=True, padx=14)
        bar = tk.Frame(self, bg=PANEL)
        bar.pack(fill="x", padx=14, pady=10)
        _button(bar, "Save charge" if existing else "Add charge", self.ok, GREEN).pack(side="right")
        _button(bar, "Cancel", self.destroy, MUTED).pack(side="right", padx=8)
        self.err = tk.Label(bar, text="", bg=PANEL, fg=RED, font=(FONT, 10), wraplength=520, justify="left")
        self.err.pack(side="left")
        self.rebuild()
        if existing:
            self.fill(existing)
        self.grab_set()

    def _row(self, label, widget, parent=None):
        r = tk.Frame(parent or self.form, bg=PANEL)
        r.pack(fill="x", pady=3)
        _label(r, label, width=30).pack(side="left")
        widget.pack(in_=r, side="left")
        widget.lift()
        return widget

    def rebuild(self):
        for w in self.form.winfo_children():
            w.destroy()
        self.err.configure(text="") if hasattr(self, "err") else None
        k = self.kind.get()
        ch = self.wiz.client_headers
        self.f = {}
        if k == "fixed":
            self.mode = tk.StringVar(value="amount")
            _radio(self.form, "Same amount for everyone", self.mode, "amount").pack(fill="x")
            self.amount = tk.StringVar()
            self._row("      Amount", _number_entry(self.form, self.amount, 14))
            _radio(self.form, "Different for each client - taken from a column of your client list", self.mode,
                   "column").pack(fill="x")
            self.f["column"] = self._row("      Column", ColMenu(self.form))
            self.f["column"].set_options(ch)
            self.desc = tk.StringVar()
            self._row("Description (optional)", _entry(self.form, self.desc, 40))
        elif k in ("metered", "extra"):
            self.fp = FilePicker(self.form, self.wiz.period, self.on_headers)
            _label(self.form, "The file that holds this month's " + ("meter readings" if k == "metered" else "extra bills") +
                   " (Excel or CSV):").pack(fill="x")
            self.fp.pack(fill="x", pady=(2, 6))
            self.f["key"] = self._row("Column with the client ID", ColMenu(self.form))
            if k == "metered":
                self.f["current_column"] = self._row("This month's reading column", ColMenu(self.form))
                self.f["previous_column"] = self._row("Last month's reading column", ColMenu(self.form))
                self.rmode = tk.StringVar(value="rate")
                r = tk.Frame(self.form, bg=PANEL)
                r.pack(fill="x", pady=(4, 0))
                _radio(r, "Price per unit", self.rmode, "rate").pack(side="left")
                self.rate = tk.StringVar()
                _number_entry(r, self.rate, 10).pack(side="left", padx=6)
                _radio(r, "or take it from column", self.rmode, "column").pack(side="left")
                self.f["rate_column"] = ColMenu(r, width=16)
                self.f["rate_column"].pack(side="left")
                _label(self.form, "Charge = (this month's reading - last month's reading) x price per unit.", MUTED, 9).pack(
                    fill="x", pady=(6, 0))
            else:
                self.f["amount_column"] = self._row("Amount column", ColMenu(self.form))
                self.f["description_column"] = self._row("Description column (optional)", ColMenu(self.form, allow_none=True))
                self.req = tk.BooleanVar(value=False)
                _check(self.form, "Every client must have an entry (otherwise no entry simply means no extra charge)",
                       self.req).pack(fill="x")
                _label(self.form, "Several rows for the same client are added together.", MUTED, 9).pack(fill="x")
        else:
            self.pct = tk.StringVar()
            self._row("Percentage", _number_entry(self.form, self.pct, 8))
            _label(self.form, "Applies to:").pack(fill="x", pady=(6, 0))
            self.scope = tk.StringVar(value="all")
            _radio(self.form, "Every charge above it in the list", self.scope, "all").pack(fill="x", padx=12)
            _radio(self.form, "Only the charges I tick:", self.scope, "some").pack(fill="x", padx=12)
            self.ticks = {}
            box = tk.Frame(self.form, bg=PANEL)
            box.pack(fill="x", padx=40)
            for n in self.earlier:
                self.ticks[n] = tk.BooleanVar()
                _check(box, n, self.ticks[n]).pack(fill="x")
            if not self.earlier:
                _label(box, "(add the other charges first - a percentage works on charges above it)", MUTED, 9).pack(fill="x")

    def fill(self, c):
        """Puts an existing charge's values back into the form."""
        k = c["kind"]
        if k == "fixed":
            if c.get("column"):
                self.mode.set("column")
                self.f["column"].select(c["column"])
            else:
                self.amount.set(str(c.get("amount", "")))
            self.desc.set(c.get("description", ""))
        elif k in ("metered", "extra"):
            self.fp.set_values(c)
            self.fp.load()
            for key in self.f:
                if c.get(key):
                    self.f[key].select(c[key])
            if k == "metered":
                if c.get("rate_column"):
                    self.rmode.set("column")
                else:
                    self.rate.set(str(c.get("rate", "")))
            else:
                self.req.set(bool(c.get("required")))
        else:
            self.pct.set(str(c.get("percent", "")))
            of = [x.strip() for x in str(c.get("of", "")).split(",") if x.strip()]
            if of:
                self.scope.set("some")
                for n in of:
                    if n in self.ticks:
                        self.ticks[n].set(True)

    def on_headers(self, headers):
        for key, menu in self.f.items():
            if key in ("key", "current_column", "previous_column", "amount_column", "description_column", "rate_column"):
                words = {"key": ("id", "house", "unit", "client", "no"), "current_column": ("current", "now", "curr", "new"),
                         "previous_column": ("previous", "prev", "last", "old"), "amount_column": ("amount", "total", "bill", "cost"),
                         "description_column": ("desc", "note"), "rate_column": ("rate", "price")}[key]
                menu.set_options(headers + (self.wiz.client_headers if key == "rate_column" else []),
                                 wizard.guess_column(headers, words))

    def _problem(self):
        """A plain message if something is missing or wrong, else ''."""
        k, name = self.kind.get(), self.name.get().strip()
        if not name:
            return "Type the name that should appear on the invoice (for example Rent)."
        if name.lower() in self.taken:
            return f"There is already a charge called '{name}'. Give this one a different name."
        if k == "fixed" and self.mode.get() == "amount" and _num(self.amount.get()) is None:
            return "Type the amount as a number, like 1500 or 1500.50."
        if k == "metered" and self.rmode.get() == "rate" and _num(self.rate.get()) is None:
            return "Type the price per unit as a number, like 150 or 0.75 - or take it from a column."
        if k == "percentage":
            p = _num(self.pct.get())
            if p is None or p <= 0:
                return "Type the percentage as a number, like 16."
            if self.scope.get() == "some" and not any(v.get() for v in self.ticks.values()):
                return "Tick at least one charge, or choose 'Every charge above it'."
        return ""

    def ok(self):
        err = self._problem()
        if err:
            self.err.configure(text=err)
            return
        k, name = self.kind.get(), self.name.get().strip()
        d = {"kind": k, "name": name}
        if k == "fixed":
            if self.mode.get() == "amount":
                d["amount"] = self.amount.get().replace(",", "").strip()
            else:
                d["column"] = self.f["column"].value()
            if self.desc.get().strip():
                d["description"] = self.desc.get().strip()
        elif k in ("metered", "extra"):
            d.update(self.fp.values())
            d["key"] = self.f["key"].value()
            if k == "metered":
                d["current_column"] = self.f["current_column"].value()
                d["previous_column"] = self.f["previous_column"].value()
                if self.rmode.get() == "rate":
                    d["rate"] = self.rate.get().replace(",", "").strip()
                else:
                    d["rate_column"] = self.f["rate_column"].value()
            else:
                d["amount_column"] = self.f["amount_column"].value()
                d["description_column"] = self.f["description_column"].value()
                d["required"] = bool(self.req.get())
        else:
            d["percent"] = self.pct.get().replace(",", "").strip()
            d["of"] = ", ".join(n for n, v in self.ticks.items() if v.get()) if self.scope.get() == "some" else ""
        try:
            wizard.build_config({"company": "x", "clients": {"file": "c.csv", "id_column": "id"}, "charges": [d],
                                 "outputs": {"pdf": True}})
        except ConfigError as e:
            self.err.configure(text=str(e))
            return
        self.result = d
        self.destroy()


def _file_shown(file, period=None):
    """'readings-{period}.csv' -> 'readings-2026-11.csv (changes each month)'."""
    name = os.path.basename(str(file))
    if period and "{" in name:
        try:
            return wizard.fill(name, period) + " (name changes each month)"
        except BillingError:
            pass
    return name


def describe_charge(c, period=None):
    k = c["kind"]
    if k == "fixed":
        return f"{c['name']}  -  " + (f"same for everyone: {c['amount']}" if c.get("amount")
                                      else f"amount from the client list column '{c['column']}'")
    if k == "metered":
        rate = c.get("rate") or f"the price in column '{c.get('rate_column')}'"
        return (f"{c['name']}  -  meter readings ('{c['current_column']}' minus '{c['previous_column']}') x {rate},"
                f"  from {_file_shown(c['file'], period)}")
    if k == "extra":
        return f"{c['name']}  -  extra bills (column '{c['amount_column']}')  from {_file_shown(c['file'], period)}"
    return f"{c['name']}  -  {c['percent']}% of " + (c.get("of") or "every charge above it")


class Wizard:
    STEPS = ["Your business", "What you bill", "Your client list", "Charges", "What to produce", "Test and save"]

    def __init__(self, master, on_done=None, config_path=None):
        sync_theme()
        self.master, self.on_done = master, on_done
        self.editing = bool(config_path)
        self.unsupported = []
        self.win = tk.Toplevel(master)
        self.win.title("BILLFLOW - " + ("change your setup" if self.editing else "new setup"))
        self.win.configure(bg=BG)
        self.win.geometry("940x720")
        self.step = 0
        self.charges, self.client_headers, self.saved_path = [], [], None
        self._cl, self._cm, self._keep = {}, {}, False
        v = tk.StringVar
        self.company, self.currency = v(), v()
        self.example_month = v(value=date.today().strftime("%Y-%m"))
        self.save_dir = v(value=os.path.join(os.path.expanduser("~"), "Documents", "billflow-setup"))
        self.out_dir = v(value="output")
        self.use_fixed, self.use_metered = tk.BooleanVar(value=True), tk.BooleanVar(value=True)
        self.use_extra, self.use_pct = tk.BooleanVar(), tk.BooleanVar()
        self.o_pdf, self.o_starter = tk.BooleanVar(value=True), tk.BooleanVar(value=True)
        self.o_ledger, self.o_json = tk.BooleanVar(value=True), tk.BooleanVar()
        self.o_template = v()
        self.test_month = v()
        if config_path:
            self.load(config_path)
        self.title = tk.Label(self.win, text="", bg=BG, fg=CYAN, font=(FONT, 20, "bold"), anchor="w")
        self.title.pack(fill="x", padx=24, pady=(14, 0))
        self.bar = tk.Label(self.win, text="", bg=BG, fg=MAG, font=(FONT, 10, "bold"), anchor="w")
        self.bar.pack(fill="x", padx=24)
        nav = tk.Frame(self.win, bg=BG)
        nav.pack(side="bottom", fill="x", padx=24, pady=(0, 14))
        self.body = tk.Frame(self.win, bg=PANEL, highlightthickness=1, highlightbackground=CYAN if gui.THEME == "neon" else BORDER)
        self.body.pack(fill="both", expand=True, padx=24, pady=10)
        self.b_back = _button(nav, "< Back", self.back, MAG)
        self.b_next = _button(nav, "Next >", self.next, GREEN)
        self.b_next.pack(side="right")
        self.b_back.pack(side="right", padx=8)
        _button(nav, "Cancel", self.win.destroy, MUTED).pack(side="right", padx=(0, 8))
        self.msg = tk.Label(nav, text="", bg=BG, fg=RED, font=(FONT, 10), anchor="w", justify="left", wraplength=560)
        self.msg.pack(side="left")
        self.show()
        if self.unsupported:
            messagebox.showwarning("BILLFLOW", "This setup has parts the wizard cannot show: " + ", ".join(self.unsupported) +
                                   ".\n\nIf you save it from the wizard, those parts will be left out. Ask us before "
                                   "changing it here.", parent=self.win)

    def load(self, config_path):
        """Fills every answer from an existing setup, so it can be changed rather than rebuilt."""
        a, self.unsupported = wizard.load_answers(config_path)
        self.company.set(a.get("company", ""))
        self.currency.set(a.get("currency", ""))
        self._terms, self._needs = a.get("payment_terms", ""), a.get("needs", "")
        if a.get("example_month"):
            self.example_month.set(a["example_month"])
        self.save_dir.set(str(os.path.dirname(os.path.abspath(config_path))))
        self.out_dir.set(a.get("output_dir") or "output")
        cl = dict(a.get("clients") or {})
        self._cl = {k: cl[k] for k in ("file", "sheet", "header_row") if cl.get(k)}
        self._cm = {k: cl.get(k, "") for k in ("id_column", "name_column", "email_column", "last_invoice_column")}
        self.charges = [dict(c) for c in a.get("charges") or []]
        kinds = {c.get("kind") for c in self.charges}
        for kind, var in (("fixed", self.use_fixed), ("metered", self.use_metered), ("extra", self.use_extra),
                          ("percentage", self.use_pct)):
            var.set(kind in kinds)
        o = a.get("outputs") or {}
        self.o_pdf.set(bool(o.get("pdf")))
        self.o_ledger.set(bool(o.get("ledger_xlsx")))
        self.o_json.set(bool(o.get("json")))
        self.o_starter.set(bool(o.get("starter_template")))
        self.o_template.set(o.get("docx_template", "") if not o.get("starter_template") else "")

    def period(self):
        p = self.example_month.get().strip()
        return p if re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", p) else date.today().strftime("%Y-%m")

    # ---- navigation ------------------------------------------------------------------------------------------
    def show(self):
        for w in self.body.winfo_children():
            w.destroy()
        self.msg.configure(text="")
        self.title.configure(text=self.STEPS[self.step])
        self.bar.configure(text=f"Step {self.step + 1} of {len(self.STEPS)}     " + "  >  ".join(
            (s.upper() if i == self.step else s) for i, s in enumerate(self.STEPS)))
        getattr(self, f"step{self.step}")()
        self.b_back.configure(state="normal" if self.step else "disabled")
        self.b_next.configure(text="Save setup" if self.step == len(self.STEPS) - 1 else "Next >")

    def back(self):
        self._stash()
        self.step = max(0, self.step - 1)
        self.show()

    def next(self):
        err = self.validate()
        if err:
            self.msg.configure(text=err, fg=RED)
            return
        if self.step == len(self.STEPS) - 1:
            self.finish()
            return
        self._stash()
        self.step += 1
        self.show()

    def _stash(self):
        """Keeps what was typed on the current page before it is closed."""
        if hasattr(self, "terms") and self.terms.winfo_exists():
            self._terms = self.terms.get("1.0", "end").strip()
        if hasattr(self, "needs") and self.needs.winfo_exists():
            self._needs = self.needs.get("1.0", "end").strip()
        if self.step == 2 and hasattr(self, "cfp") and self.cfp.winfo_exists():
            self._cl = self.cfp.values()
            self._cm = {"id_column": self.id_menu.value(), "name_column": self.name_menu.value(),
                        "email_column": self.email_menu.value(), "last_invoice_column": self.last_menu.value()}

    def validate(self):
        if self.step == 0 and not self.company.get().strip():
            return "Type your company name."
        if self.step == 2:
            if not self.cfp.file.get().strip():
                return "Choose the file that lists your clients."
            if not self.cfp.load():
                return "The client file could not be read - see the message above."
            if not self.id_menu.value():
                return "Choose which column holds each client's ID."
        if self.step == 3 and not self.charges:
            return "Add at least one charge (press + Add a charge)."
        if self.step == 4:
            if not (self.o_pdf.get() or self.o_starter.get() or self.o_template.get().strip() or self.o_ledger.get() or self.o_json.get()):
                return "Tick at least one thing to produce."
        return ""

    def _p(self, text, color=None, size=10, **kw):
        _label(self.body, text, color, size, **kw).pack(fill="x", padx=18, pady=(8, 0))

    def _hint(self, text, **kw):
        _label(self.body, text, MUTED, 9, wraplength=840).pack(fill="x", padx=18, **kw)

    # ---- steps -----------------------------------------------------------------------------------------------
    def step0(self):
        self._p("Tell us about your business. This is printed on your invoices.", CYAN, 11)
        for text, var, w in (("Company name", self.company, 40), ("Currency (e.g. KES, USD)", self.currency, 10)):
            r = tk.Frame(self.body, bg=PANEL)
            r.pack(fill="x", padx=18, pady=5)
            _label(r, text, width=30).pack(side="left")
            _entry(r, var, w).pack(side="left")
        _label(self.body, "Payment details (printed at the bottom of each invoice - bank, M-Pesa paybill, due date...)").pack(
            fill="x", padx=18, pady=(6, 0))
        self.terms = tk.Text(self.body, height=4, font=(FONT, 10), bg=INPUT, fg=INPUT_FG, insertbackground=INPUT_FG,
                             relief="flat", highlightthickness=1, highlightbackground=BORDER, wrap="word")
        self.terms.pack(fill="x", padx=18, pady=4)
        if getattr(self, "_terms", ""):
            self.terms.insert("1.0", self._terms)
        r = tk.Frame(self.body, bg=PANEL)
        r.pack(fill="x", padx=18, pady=(10, 0))
        _label(r, "Month of the files you will pick", width=30).pack(side="left")
        gui.MonthMenu(r, self.example_month, 18).pack(side="left")
        self._hint("If your file names contain the month (for example 'Readings November 2026.xlsx'), choose that month here. "
                   "BILLFLOW then looks for next month's file by itself.", pady=(2, 0))
        r = tk.Frame(self.body, bg=PANEL)
        r.pack(fill="x", padx=18, pady=(12, 0))
        _label(r, "Save this setup in the folder", width=30).pack(side="left")
        _entry(r, self.save_dir, 44).pack(side="left")
        _button(r, "Choose folder...", lambda: self.save_dir.set(
            os.path.normpath(filedialog.askdirectory(parent=self.win) or self.save_dir.get()))).pack(side="left", padx=6)
        if self.editing:
            self._hint("You are changing an existing setup - saving replaces it in this folder.", pady=(4, 0))

    def step1(self):
        self._p("What do you bill your clients for? Tick everything that applies.", CYAN, 11)
        for text, var in (("A fixed amount per client (rent, service charge)", self.use_fixed),
                          ("Usage worked out from meter readings (water, power)", self.use_metered),
                          ("Extra bills that get added to a client (e.g. electricity bills in separate files)", self.use_extra),
                          ("A percentage on top (VAT, late fees)", self.use_pct)):
            _check(self.body, text, var).pack(fill="x", padx=24, pady=1)
        self._hint("Your ticks are a checklist for the Charges step - you will add each charge there.", pady=(4, 0))
        self._p("Notes for yourself (optional)", TEXT, 10, bold=True)
        self._hint("Only kept as a reminder in the setup. It does not change how invoices are worked out. Example: 'Each month "
                   "I bill 40 tenants rent plus water from the meter log; invoice numbers continue from the last one.'")
        self.needs = tk.Text(self.body, height=5, font=(FONT, 10), bg=INPUT, fg=INPUT_FG, insertbackground=INPUT_FG,
                             relief="flat", highlightthickness=1, highlightbackground=BORDER, wrap="word")
        self.needs.pack(fill="x", padx=18, pady=6)
        self.needs.insert("1.0", getattr(self, "_needs", ""))
        self._hint("Need something unusual that none of these cover? Write it above and ask us - it can be added as a custom rule.")

    def step2(self):
        self._p("Choose the file that lists your clients (Excel or CSV) - one row per client.", CYAN, 11)
        self.cfp = FilePicker(self.body, self.period, self.on_client_headers)
        self.cfp.pack(fill="x", padx=18, pady=4)
        if self._cl:
            self.cfp.set_values(self._cl)
        self.id_menu, self.name_menu = ColMenu(self.body), ColMenu(self.body)
        self.email_menu, self.last_menu = ColMenu(self.body, True), ColMenu(self.body, True)
        for text, m in (("Column with each client's ID (house / unit no.)", self.id_menu),
                        ("Column with the client's name", self.name_menu),
                        ("Column with their email (optional)", self.email_menu),
                        ("Column with their last invoice number (optional)", self.last_menu)):
            r = tk.Frame(self.body, bg=PANEL)
            r.pack(fill="x", padx=18, pady=3)
            _label(r, text, width=48).pack(side="left")
            m.pack(in_=r, side="left")
            m.lift()
        self._hint("Last invoice number: only pick a column if your client list keeps each client's most recent invoice number. "
                   "Example: if H1's last invoice was 12, this month's will be H1/13. No such column? Leave it as (none) and "
                   "numbering starts at 1 (H1/1). After that, BILLFLOW keeps count by itself.", pady=(8, 0))
        for k, m in (("id_column", self.id_menu), ("name_column", self.name_menu), ("email_column", self.email_menu),
                     ("last_invoice_column", self.last_menu)):
            if self._cm.get(k):
                m.select(self._cm[k])
        if self.client_headers:
            self.on_client_headers(self.client_headers, keep=bool(self._cm))
        elif self._cl.get("file"):
            self._keep = bool(self._cm)
            try:
                self.cfp.load()
            finally:
                self._keep = False

    def on_client_headers(self, headers, keep=False):
        keep = keep or self._keep
        self.client_headers = list(headers)
        g = (lambda words: "") if keep else (lambda words: wizard.guess_column(headers, words))
        self.id_menu.set_options(headers, g(("id", "house", "unit", "no", "code")))
        self.name_menu.set_options(headers, g(("name", "client", "tenant", "customer")))
        self.email_menu.set_options(headers, g(("email", "mail")) or None)
        self.last_menu.set_options(headers, g(("last", "invoice")) or None)

    def step3(self):
        self._p("Add one line per charge, in the order they should appear on the invoice.", CYAN, 11)
        frame = tk.Frame(self.body, bg=PANEL)
        frame.pack(fill="both", expand=True, padx=18, pady=6)
        self.lst = tk.Listbox(frame, font=(FONT, 10), bg=INPUT, fg=TEXT, selectbackground=MAG, selectforeground=HOVER_FG,
                              relief="flat", highlightthickness=1, highlightbackground=BORDER, height=9)
        self.lst.pack(fill="both", expand=True)
        self.lst.bind("<Double-Button-1>", lambda e: self.edit_charge())
        bar = tk.Frame(self.body, bg=PANEL)
        bar.pack(fill="x", padx=18, pady=6)
        _button(bar, "+ Add a charge", self.add_charge, GREEN).pack(side="left")
        _button(bar, "Change selected", self.edit_charge, CYAN).pack(side="left", padx=8)
        _button(bar, "Remove selected", self.remove_charge, RED).pack(side="left")
        _button(bar, "Move up", lambda: self.move(-1), YEL).pack(side="left", padx=8)
        _button(bar, "Move down", lambda: self.move(1), YEL).pack(side="left")
        self.todo = tk.StringVar()
        tk.Label(self.body, textvariable=self.todo, bg=PANEL, fg=YEL, font=(FONT, 10), anchor="w", justify="left",
                 wraplength=840).pack(fill="x", padx=18)
        self.refresh_list()

    def refresh_list(self):
        self.lst.delete(0, "end")
        for c in self.charges:
            self.lst.insert("end", describe_charge(c, self.period()))
        have = {c["kind"] for c in self.charges}
        missing = [t for k, t, v in (("fixed", "a fixed amount", self.use_fixed), ("metered", "meter readings", self.use_metered),
                                     ("extra", "extra bills", self.use_extra), ("percentage", "a percentage", self.use_pct))
                   if v.get() and k not in have]
        self.todo.set(("You ticked " + ", ".join(missing) + " but have not added it yet.") if missing else
                      ("" if self.charges else "Nothing added yet - press + Add a charge."))

    def add_charge(self):
        d = ChargeDialog(self)
        self.win.wait_window(d)
        if d.result:
            self.charges.append(d.result)
            self.refresh_list()

    def edit_charge(self):
        s = self.lst.curselection()
        if not s:
            self.msg.configure(text="Click a charge in the list first, then press Change selected.", fg=RED)
            return
        i = s[0]
        d = ChargeDialog(self, self.charges[i], earlier=[c["name"] for c in self.charges[:i]])
        self.win.wait_window(d)
        if d.result:
            self.charges[i] = d.result
            self.refresh_list()
            self.lst.selection_set(i)

    def remove_charge(self):
        s = self.lst.curselection()
        if s and messagebox.askyesno("BILLFLOW", f"Remove '{self.charges[s[0]]['name']}'?", parent=self.win):
            del self.charges[s[0]]
            self.refresh_list()

    def move(self, delta):
        s = self.lst.curselection()
        if s and 0 <= s[0] + delta < len(self.charges):
            i = s[0]
            self.charges[i], self.charges[i + delta] = self.charges[i + delta], self.charges[i]
            self.refresh_list()
            self.lst.selection_set(i + delta)

    def step4(self):
        self._p("What should BILLFLOW make for each client every month?", CYAN, 11)
        for text, var in (("A PDF invoice", self.o_pdf),
                          ("A Word invoice, using a ready-made template", self.o_starter),
                          ("A client ledger in Excel (running balance; you type in payments)", self.o_ledger),
                          ("A data file (JSON) for other systems - only if someone asked you for it", self.o_json)):
            _check(self.body, text, var).pack(fill="x", padx=24, pady=2)
        _label(self.body, "Have your own Word invoice design? Choose it here instead of the ready-made one:").pack(
            fill="x", padx=18, pady=(12, 2))
        r = tk.Frame(self.body, bg=PANEL)
        r.pack(fill="x", padx=18)
        _entry(r, self.o_template, 58).pack(side="left")
        _button(r, "Choose template...", self.choose_template).pack(side="left", padx=6)
        self._hint("In your template, type {{client_name}}, {{invoice_no}}, {{date}}, {{period}}, {{total}}, {{payment_terms}} "
                   "where those should go, and one table row with {{item}} {{description}} {{qty}} {{unit_price}} {{amount}}.")
        r = tk.Frame(self.body, bg=PANEL)
        r.pack(fill="x", padx=18, pady=(14, 0))
        _label(r, "Folder for the results", width=26).pack(side="left")
        _entry(r, self.out_dir, 44).pack(side="left")
        _button(r, "Choose folder...", lambda: self.out_dir.set(
            os.path.normpath(filedialog.askdirectory(parent=self.win) or self.out_dir.get()))).pack(side="left", padx=6)
        self._hint("Each month's results go in a folder named for the month. Your own files are never changed.", pady=(4, 0))

    def choose_template(self):
        p = filedialog.askopenfilename(parent=self.win, filetypes=[("Word", "*.docx")])
        if p:
            self.o_template.set(os.path.normpath(p))

    def step5(self):
        self._p("Test your setup on a month before saving. Testing never saves or changes anything.", CYAN, 11)
        r = tk.Frame(self.body, bg=PANEL)
        r.pack(fill="x", padx=18, pady=4)
        _label(r, "Month to test", width=16).pack(side="left")
        self.test_month.set(self.test_month.get() or self.period())
        gui.MonthMenu(r, self.test_month, 18).pack(side="left")
        _button(r, "Run test", self.run_test, GREEN).pack(side="left", padx=10)
        self.out = tk.Text(self.body, font=(FONT, 10), bg=INPUT, fg=TEXT, relief="flat", wrap="word", height=20,
                           highlightthickness=1, highlightbackground=BORDER, state="disabled", padx=8, pady=6)
        self.out.pack(fill="both", expand=True, padx=18, pady=6)
        for tag, col in (("bad", RED), ("good", GREEN), ("info", TEXT), ("head", CYAN), ("muted", MUTED)):
            self.out.tag_configure(tag, foreground=col)
        self.out.tag_configure("head", font=(FONT, 11, "bold"))
        self.out.tag_configure("fix", lmargin1=24, lmargin2=24, foreground=TEXT)
        self.say(f"Company: {self.company.get()}", "info")
        self.say(f"Client list: {os.path.basename(self._cl.get('file', ''))}", "info")
        for c in self.charges:
            self.say("  + " + describe_charge(c, self.period()), "good")
        self.say(f"\nThe setup will be saved in: {self.save_dir.get()}", "muted")
        self.say("Press Run test to read your files for the chosen month.", "info")

    def say(self, text, tag="good"):
        self.out.configure(state="normal")
        self.out.insert("end", text + "\n", tag)
        self.out.configure(state="disabled")

    # ---- answers ---------------------------------------------------------------------------------------------
    def collect(self):
        self._stash()
        cl = dict(self._cl)
        cl.update({k: v for k, v in self._cm.items() if v})
        outs = {"pdf": self.o_pdf.get(), "ledger_xlsx": self.o_ledger.get(), "json": self.o_json.get()}
        if self.o_template.get().strip():
            outs["docx_template"] = self.o_template.get().strip()
        elif self.o_starter.get():
            outs["starter_template"] = True
        return {"company": self.company.get().strip(), "currency": self.currency.get().strip(),
                "payment_terms": getattr(self, "_terms", ""), "needs": getattr(self, "_needs", ""), "clients": cl,
                "charges": list(self.charges), "outputs": outs, "output_dir": self.out_dir.get().strip() or "output",
                "example_month": self.period()}

    @staticmethod
    def _plain_problem(p):
        from .friendly import explain
        label, _, rest = p.partition(": ")
        return f"{label}: {explain(rest).text}" if rest else p

    def run_test(self):
        from .engine import plan
        from .config import load_config
        import tempfile
        self.out.configure(state="normal")
        self.out.delete("1.0", "end")
        self.out.configure(state="disabled")
        a, period = self.collect(), self.test_month.get().strip()
        problems = wizard.check_answers(a, period)
        if problems:
            self.say(f"Some files could not be read for {month_label(period)}:", "bad")
            for p in problems:
                self.say("  • " + self._plain_problem(p), "bad")
            self.say("\nGo Back to fix the setup, or test a month whose files you already have.", "info")
            return
        try:
            with tempfile.TemporaryDirectory() as tmp:
                cfg = load_config(wizard.write_config(a, tmp))
                results = plan(cfg, period)
                items = describe_results(results, cfg, period, a.get("currency", ""))
        except (ConfigError, BillingError, OSError) as e:
            self.say("Problem: " + self._plain_problem(str(e)), "bad")
            return
        ok = [r for r in results if r.status == "OK"]
        bad = sum(r.status == "ERROR" for r in results)
        self.say(f"{month_label(period)}: {gui.plural(len(ok), 'invoice')} ready, "
                 f"{gui.plural(bad, 'client')} {'needs' if bad == 1 else 'need'} attention, "
                 f"total {sum(r.invoice.total for r in ok):,.2f}\n", "head")
        for d in items:
            tag = "good" if d["status"] == "OK" else "bad"
            self.say(f"• {d['client_id']}  {d['client_name']}   {d['text']}", tag)
            if d["fix"]:
                self.say("What to do: " + d["fix"], "fix")
        self.say("\nNothing was saved. If this looks right, press Save setup.", "muted")

    def finish(self):
        a = self.collect()
        problems = wizard.check_answers(a, self.test_month.get().strip() or self.period())
        if problems and not messagebox.askyesno("BILLFLOW", "Some files could not be read for the test month:\n\n" +
                                                "\n".join(self._plain_problem(p) for p in problems[:6]) +
                                                "\n\nSave the setup anyway?", parent=self.win):
            return
        if self.unsupported and not messagebox.askyesno(
                "BILLFLOW", "Saving here will leave out: " + ", ".join(self.unsupported) + ".\n\nSave anyway?", parent=self.win):
            return
        try:
            path = wizard.write_config(a, self.save_dir.get().strip())
        except (ConfigError, OSError) as e:
            self.msg.configure(text=str(e), fg=RED)
            return
        self.saved_path = path
        messagebox.showinfo("BILLFLOW", f"Setup saved.\n\nBack on the main screen, choose the month and press "
                                        f"Check invoices.\n\n(Saved in {path.parent})", parent=self.win)
        if self.on_done:
            self.on_done(str(path), self.test_month.get().strip() or self.period())
        self.win.destroy()


def open_wizard(master, on_done=None, config_path=None):
    return Wizard(master, on_done, config_path)
