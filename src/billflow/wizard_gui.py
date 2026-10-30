"""BILLFLOW setup wizard screen: a business describes what it needs, links its own files and folders,
and the wizard writes the config. Nothing here touches the linked files except to read their column names."""
import os
import re
import tkinter as tk
from datetime import date
from tkinter import filedialog, messagebox

from . import wizard
from .config import ConfigError
from .gui import BG, CYAN, FONT, GREEN, MAG, PANEL, RED, YEL
from .models import BillingError

NONE = "(none)"
INPUT = "#02030a"


def _label(parent, text, color=GREEN, size=10, bold=False, **kw):
    return tk.Label(parent, text=text, bg=PANEL, fg=color, font=(FONT, size, "bold" if bold else "normal"), anchor="w",
                    justify="left", **kw)


def _entry(parent, var, width=30):
    return tk.Entry(parent, textvariable=var, width=width, font=(FONT, 11), bg=INPUT, fg=CYAN, insertbackground=CYAN,
                    relief="flat", highlightthickness=1, highlightbackground="#0b5560", highlightcolor=CYAN)


def _button(parent, text, cmd, color=CYAN, width=None):
    b = tk.Button(parent, text=text, command=cmd, font=(FONT, 10, "bold"), fg=color, bg=PANEL, bd=0, relief="flat",
                  activebackground=color, activeforeground=BG, cursor="hand2", highlightthickness=1,
                  highlightbackground=color, highlightcolor=color, padx=10, pady=3, width=width)
    b.bind("<Enter>", lambda e: b.configure(bg=color, fg=BG))
    b.bind("<Leave>", lambda e: b.configure(bg=PANEL, fg=color))
    return b


def _check(parent, text, var):
    return tk.Checkbutton(parent, text=text, variable=var, bg=PANEL, fg=CYAN, selectcolor=INPUT, activebackground=PANEL,
                          activeforeground=CYAN, font=(FONT, 10), anchor="w", highlightthickness=0, bd=0)


def _radio(parent, text, var, value, cmd=None):
    return tk.Radiobutton(parent, text=text, variable=var, value=value, command=cmd, bg=PANEL, fg=CYAN, selectcolor=INPUT,
                          activebackground=PANEL, activeforeground=CYAN, font=(FONT, 10), anchor="w", highlightthickness=0, bd=0)


class ColMenu(tk.Frame):
    """A drop-down of column names found in the linked file."""

    def __init__(self, master, allow_none=False, width=24):
        super().__init__(master, bg=PANEL)
        self.var = tk.StringVar()
        self.allow_none = allow_none
        self.menu = tk.OptionMenu(self, self.var, "")
        self.menu.configure(font=(FONT, 10), bg=INPUT, fg=CYAN, activebackground=CYAN, activeforeground=BG, relief="flat",
                            highlightthickness=1, highlightbackground="#0b5560", width=width, anchor="w", bd=0)
        self.menu["menu"].configure(bg=INPUT, fg=CYAN, font=(FONT, 10), activebackground=CYAN, activeforeground=BG)
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
        _button(top, "BROWSE", self.browse).pack(side="left")
        row = tk.Frame(self, bg=PANEL)
        row.pack(fill="x", pady=(4, 0))
        if with_sheet:
            _label(row, "Excel sheet").pack(side="left")
            _entry(row, self.sheet, 18).pack(side="left", padx=(6, 12))
            _label(row, "Header row").pack(side="left")
            _entry(row, self.hrow, 3).pack(side="left", padx=(6, 12))
        _button(row, "READ COLUMNS", self.load, GREEN).pack(side="left")
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
            self.status.set("Problem: " + str(e))
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
    """Add one charge. Asks only what that kind of charge needs."""

    def __init__(self, wiz):
        super().__init__(wiz.win)
        self.wiz, self.result = wiz, None
        self.title("Add a charge")
        self.configure(bg=PANEL)
        self.geometry("780x560")
        self.transient(wiz.win)
        self.kind = tk.StringVar(value="fixed")
        self.name = tk.StringVar()
        tk.Frame(self, bg=PANEL, height=6).pack()
        _label(self, "  What kind of charge is it?", MAG, 11, True).pack(fill="x")
        kinds = tk.Frame(self, bg=PANEL)
        kinds.pack(fill="x", padx=14)
        for text, val in (("Fixed amount (rent, service charge)", "fixed"),
                          ("Usage from meter readings (water, power)", "metered"),
                          ("Extra bills added to a client (e.g. electricity bills)", "extra"),
                          ("Percentage of other charges (VAT, late fee)", "percentage")):
            _radio(kinds, text, self.kind, val, self.rebuild).pack(fill="x")
        nm = tk.Frame(self, bg=PANEL)
        nm.pack(fill="x", padx=14, pady=8)
        _label(nm, "Name on the invoice  ").pack(side="left")
        _entry(nm, self.name, 36).pack(side="left")
        self.form = tk.Frame(self, bg=PANEL)
        self.form.pack(fill="both", expand=True, padx=14)
        bar = tk.Frame(self, bg=PANEL)
        bar.pack(fill="x", padx=14, pady=10)
        _button(bar, "ADD CHARGE", self.ok, GREEN).pack(side="right")
        _button(bar, "CANCEL", self.destroy, RED).pack(side="right", padx=8)
        self.err = tk.Label(bar, text="", bg=PANEL, fg=RED, font=(FONT, 9), wraplength=480, justify="left")
        self.err.pack(side="left")
        self.rebuild()
        self.grab_set()

    def _row(self, label, widget):
        r = tk.Frame(self.form, bg=PANEL)
        r.pack(fill="x", pady=3)
        _label(r, label, width=26).pack(side="left")
        widget.pack(in_=r, side="left")
        widget.lift()
        return widget

    def rebuild(self):
        for w in self.form.winfo_children():
            w.destroy()
        k = self.kind.get()
        ch = self.wiz.client_headers
        self.f = {}
        if k == "fixed":
            self.mode = tk.StringVar(value="amount")
            _radio(self.form, "Same amount for everyone", self.mode, "amount").pack(fill="x")
            self.amount = tk.StringVar()
            self._row("   Amount", _entry(self.form, self.amount, 14))
            _radio(self.form, "Different per client - taken from a column of the client list", self.mode, "column").pack(fill="x")
            self.f["column"] = self._row("   Column", ColMenu(self.form))
            self.f["column"].set_options(ch)
            self.desc = tk.StringVar()
            self._row("Description (optional)", _entry(self.form, self.desc, 40))
        elif k in ("metered", "extra"):
            self.fp = FilePicker(self.form, self.wiz.period, self.on_headers)
            _label(self.form, "The file that holds this month's " + ("readings" if k == "metered" else "bills") +
                   " (the folder is part of the path):").pack(fill="x")
            self.fp.pack(fill="x", pady=(2, 6))
            self.f["key"] = self._row("Column with the client ID", ColMenu(self.form))
            if k == "metered":
                self.f["current_column"] = self._row("Current reading column", ColMenu(self.form))
                self.f["previous_column"] = self._row("Previous reading column", ColMenu(self.form))
                self.rmode = tk.StringVar(value="rate")
                r = tk.Frame(self.form, bg=PANEL)
                r.pack(fill="x")
                _radio(r, "Price per unit", self.rmode, "rate").pack(side="left")
                self.rate = tk.StringVar()
                _entry(r, self.rate, 10).pack(side="left", padx=6)
                _radio(r, "or from column", self.rmode, "column").pack(side="left")
                self.f["rate_column"] = ColMenu(r, width=16)
                self.f["rate_column"].pack(side="left")
            else:
                self.f["amount_column"] = self._row("Amount column", ColMenu(self.form))
                self.f["description_column"] = self._row("Description column (optional)", ColMenu(self.form, allow_none=True))
                self.req = tk.BooleanVar(value=False)
                _check(self.form, "Every client must have an entry (otherwise a missing entry just means no extra charge)",
                       self.req).pack(fill="x")
                _label(self.form, "Several rows for the same client are added together.", "#3a6a70", 9).pack(fill="x")
        else:
            self.pct = tk.StringVar()
            self._row("Percentage", _entry(self.form, self.pct, 10))
            self.of = tk.StringVar()
            self._row("Applies to (names, comma-separated)", _entry(self.form, self.of, 36))
            _label(self.form, "Leave 'Applies to' empty to apply it to every charge above it in the list.", "#3a6a70", 9).pack(fill="x")

    def on_headers(self, headers):
        for key, menu in self.f.items():
            if key in ("key", "current_column", "previous_column", "amount_column", "description_column", "rate_column"):
                words = {"key": ("id", "house", "unit", "client", "no"), "current_column": ("current", "now", "curr", "new"),
                         "previous_column": ("previous", "prev", "last", "old"), "amount_column": ("amount", "total", "bill", "cost"),
                         "description_column": ("desc", "note"), "rate_column": ("rate", "price")}[key]
                menu.set_options(headers + (self.wiz.client_headers if key == "rate_column" else []),
                                 wizard.guess_column(headers, words))

    def ok(self):
        k, name = self.kind.get(), self.name.get().strip()
        d = {"kind": k, "name": name}
        if k == "fixed":
            if self.mode.get() == "amount":
                d["amount"] = self.amount.get().strip()
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
                    d["rate"] = self.rate.get().strip()
                else:
                    d["rate_column"] = self.f["rate_column"].value()
            else:
                d["amount_column"] = self.f["amount_column"].value()
                d["description_column"] = self.f["description_column"].value()
                d["required"] = bool(self.req.get())
        else:
            d["percent"] = self.pct.get().strip()
            d["of"] = self.of.get().strip()
        try:
            wizard.build_config({"company": "x", "clients": {"file": "c.csv", "id_column": "id"}, "charges": [d],
                                 "outputs": {"pdf": True}})
        except ConfigError as e:
            self.err.configure(text=str(e))
            return
        self.result = d
        self.destroy()


def describe_charge(c):
    k = c["kind"]
    if k == "fixed":
        return f"{c['name']}  -  fixed, " + (f"amount {c['amount']}" if c.get("amount") else f"from column '{c['column']}'")
    if k == "metered":
        rate = c.get("rate") or f"column '{c.get('rate_column')}'"
        return f"{c['name']}  -  usage: ({c['current_column']} - {c['previous_column']}) x {rate}   [{os.path.basename(c['file'])}]"
    if k == "extra":
        return f"{c['name']}  -  extra bills from column '{c['amount_column']}'   [{os.path.basename(c['file'])}]"
    return f"{c['name']}  -  {c['percent']}% of " + (c.get("of") or "all charges above")


class Wizard:
    STEPS = ["YOUR BUSINESS", "WHAT YOU NEED", "CLIENT LIST", "CHARGES", "WHAT TO PRODUCE", "CHECK & SAVE"]

    def __init__(self, master, on_done=None):
        self.master, self.on_done = master, on_done
        self.win = tk.Toplevel(master)
        self.win.title("BILLFLOW // SETUP WIZARD")
        self.win.configure(bg=BG)
        self.win.geometry("900x700")
        self.step = 0
        self.charges, self.client_headers, self.saved_path = [], [], None
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
        self.title = tk.Label(self.win, text="", bg=BG, fg=CYAN, font=(FONT, 20, "bold"), anchor="w")
        self.title.pack(fill="x", padx=24, pady=(14, 0))
        self.bar = tk.Label(self.win, text="", bg=BG, fg=MAG, font=(FONT, 9, "bold"), anchor="w")
        self.bar.pack(fill="x", padx=24)
        nav = tk.Frame(self.win, bg=BG)
        nav.pack(side="bottom", fill="x", padx=24, pady=(0, 14))
        self.body = tk.Frame(self.win, bg=PANEL, highlightthickness=1, highlightbackground=CYAN)
        self.body.pack(fill="both", expand=True, padx=24, pady=10)
        self.b_back = _button(nav, "< BACK", self.back, MAG)
        self.b_next = _button(nav, "NEXT >", self.next, GREEN)
        self.b_next.pack(side="right")
        self.b_back.pack(side="right", padx=8)
        self.msg = tk.Label(nav, text="", bg=BG, fg=RED, font=(FONT, 9), anchor="w", justify="left", wraplength=560)
        self.msg.pack(side="left")
        self.show()

    def period(self):
        p = self.example_month.get().strip()
        return p if re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", p) else date.today().strftime("%Y-%m")

    # ---- navigation ------------------------------------------------------------------------------------------
    def show(self):
        for w in self.body.winfo_children():
            w.destroy()
        self.msg.configure(text="")
        self.title.configure(text=self.STEPS[self.step])
        self.bar.configure(text=f"STEP {self.step + 1} OF {len(self.STEPS)}   " + "  ".join(
            ("[#]" if i <= self.step else "[ ]") for i in range(len(self.STEPS))))
        getattr(self, f"step{self.step}")()
        self.b_back.configure(state="normal" if self.step else "disabled")
        self.b_next.configure(text="SAVE SETUP" if self.step == len(self.STEPS) - 1 else "NEXT >")

    def back(self):
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
        self.step += 1
        self.show()

    def validate(self):
        if self.step == 0 and not self.company.get().strip():
            return "Enter your company name."
        if self.step == 0 and not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", self.example_month.get().strip()):
            return "Month must look like 2026-11."
        if self.step == 2:
            if not self.cfp.file.get().strip():
                return "Choose the file that lists your clients."
            if not self.cfp.load():
                return "The client file could not be read - see the message above."
            if not self.id_menu.value():
                return "Choose which column holds each client's ID."
        if self.step == 3 and not self.charges:
            return "Add at least one charge."
        if self.step == 4:
            if not (self.o_pdf.get() or self.o_starter.get() or self.o_template.get().strip() or self.o_ledger.get() or self.o_json.get()):
                return "Choose at least one thing to produce."
        return ""

    def _p(self, text, color=GREEN, size=10, **kw):
        _label(self.body, text, color, size, **kw).pack(fill="x", padx=18, pady=(8, 0))

    # ---- steps -----------------------------------------------------------------------------------------------
    def step0(self):
        self._p("Tell us about your business. This appears on your invoices.", CYAN, 11)
        for text, var, w in (("Company name", self.company, 40), ("Currency (e.g. KES, USD)", self.currency, 10)):
            r = tk.Frame(self.body, bg=PANEL)
            r.pack(fill="x", padx=18, pady=5)
            _label(r, text, width=26).pack(side="left")
            _entry(r, var, w).pack(side="left")
        _label(self.body, "Payment details / terms (printed at the bottom of each invoice)").pack(fill="x", padx=18, pady=(6, 0))
        self.terms = tk.Text(self.body, height=4, font=(FONT, 10), bg=INPUT, fg=CYAN, insertbackground=CYAN, relief="flat",
                             highlightthickness=1, highlightbackground="#0b5560")
        self.terms.pack(fill="x", padx=18, pady=4)
        if getattr(self, "_terms", ""):
            self.terms.insert("1.0", self._terms)
        r = tk.Frame(self.body, bg=PANEL)
        r.pack(fill="x", padx=18, pady=5)
        _label(r, "Month of your example files", width=30).pack(side="left")
        _entry(r, self.example_month, 10).pack(side="left")
        _label(r, "  YYYY-MM (month names in file names become patterns)", "#3a6a70", 9).pack(side="left")
        r = tk.Frame(self.body, bg=PANEL)
        r.pack(fill="x", padx=18, pady=5)
        _label(r, "Save this setup in folder", width=26).pack(side="left")
        _entry(r, self.save_dir, 46).pack(side="left")
        _button(r, "BROWSE", lambda: self.save_dir.set(filedialog.askdirectory() or self.save_dir.get())).pack(side="left", padx=6)

    def step1(self):
        if hasattr(self, "terms") and self.terms.winfo_exists():
            self._terms = self.terms.get("1.0", "end").strip()
        self._p("In your own words: what do you bill, and what should the system do each month?", CYAN, 11)
        _label(self.body, "Example: 'Each month I bill 40 tenants rent plus water from the meter log. Some also get an electricity "
                          "bill. Invoice numbers continue from the last one. I send a PDF to each tenant.'", "#3a6a70", 9,
               wraplength=800).pack(fill="x", padx=18)
        self.needs = tk.Text(self.body, height=6, font=(FONT, 10), bg=INPUT, fg=CYAN, insertbackground=CYAN, relief="flat",
                             highlightthickness=1, highlightbackground="#0b5560", wrap="word")
        self.needs.pack(fill="x", padx=18, pady=6)
        self.needs.insert("1.0", getattr(self, "_needs", ""))
        _label(self.body, "Your note is saved with the setup as a record. The questions below decide what gets built, "
                          "so tick everything that applies:", GREEN, 9, wraplength=800).pack(fill="x", padx=18, pady=(6, 2))
        for text, var in (("A fixed amount per client (rent, service charge)", self.use_fixed),
                          ("Usage worked out from meter readings (water, power)", self.use_metered),
                          ("Extra bills that get added to a client (e.g. electricity bills in separate files)", self.use_extra),
                          ("A percentage on top (VAT, late fees)", self.use_pct)):
            _check(self.body, text, var).pack(fill="x", padx=24)
        _label(self.body, "Need something unusual? Describe it above - it can be added as a custom rule (see the guide).",
               "#3a6a70", 9, wraplength=800).pack(fill="x", padx=18, pady=(8, 0))

    def step2(self):
        if hasattr(self, "needs") and self.needs.winfo_exists():
            self._needs = self.needs.get("1.0", "end").strip()
        self._p("Link the file that lists your clients (Excel or CSV). One row per client.", CYAN, 11)
        self.cfp = FilePicker(self.body, self.period, self.on_client_headers)
        self.cfp.pack(fill="x", padx=18, pady=4)
        if getattr(self, "_cl", None):
            self.cfp.set_values(self._cl)
        self.id_menu, self.name_menu = ColMenu(self.body), ColMenu(self.body)
        self.email_menu, self.last_menu = ColMenu(self.body, True), ColMenu(self.body, True)
        for text, m in (("Column with each client's ID", self.id_menu), ("Column with the client's name", self.name_menu),
                        ("Column with their email (optional)", self.email_menu),
                        ("Column with their last invoice number (optional)", self.last_menu)):
            r = tk.Frame(self.body, bg=PANEL)
            r.pack(fill="x", padx=18, pady=3)
            _label(r, text, width=46).pack(side="left")
            m.pack(side="left")
        _label(self.body, "New invoice numbers are the last number + 1, written like ID/number (e.g. H1/13). If you have no 'last "
                          "invoice' column, numbering starts at 1.", "#3a6a70", 9, wraplength=800).pack(fill="x", padx=18, pady=(6, 0))
        if self.client_headers:
            self.on_client_headers(self.client_headers, keep=True)

    def on_client_headers(self, headers, keep=False):
        self.client_headers = list(headers)
        self.id_menu.set_options(headers, wizard.guess_column(headers, ("id", "house", "unit", "no", "code")))
        self.name_menu.set_options(headers, wizard.guess_column(headers, ("name", "client", "tenant", "customer")))
        self.email_menu.set_options(headers, wizard.guess_column(headers, ("email", "mail")) or NONE)
        self.last_menu.set_options(headers, wizard.guess_column(headers, ("last", "invoice")) or NONE)

    def step3(self):
        self._cl = self.cfp.values()
        self._cm = {"id_column": self.id_menu.value(), "name_column": self.name_menu.value(),
                    "email_column": self.email_menu.value(), "last_invoice_column": self.last_menu.value()}
        self._p("What do you bill? Add one line per charge, in the order they should appear.", CYAN, 11)
        wants = [t for t, v in (("fixed amounts", self.use_fixed), ("meter-reading usage", self.use_metered),
                                ("extra bills", self.use_extra), ("a percentage", self.use_pct)) if v.get()]
        if wants:
            _label(self.body, "You said you need: " + ", ".join(wants) + ".", "#3a6a70", 9).pack(fill="x", padx=18)
        frame = tk.Frame(self.body, bg=PANEL)
        frame.pack(fill="both", expand=True, padx=18, pady=6)
        self.lst = tk.Listbox(frame, font=(FONT, 10), bg=INPUT, fg=GREEN, selectbackground=MAG, selectforeground=BG,
                              relief="flat", highlightthickness=1, highlightbackground="#0b5560", height=10)
        self.lst.pack(fill="both", expand=True)
        bar = tk.Frame(self.body, bg=PANEL)
        bar.pack(fill="x", padx=18, pady=6)
        _button(bar, "+ ADD CHARGE", self.add_charge, GREEN).pack(side="left")
        _button(bar, "REMOVE SELECTED", self.remove_charge, RED).pack(side="left", padx=8)
        _button(bar, "MOVE UP", lambda: self.move(-1), YEL).pack(side="left")
        self.refresh_list()

    def refresh_list(self):
        self.lst.delete(0, "end")
        for c in self.charges:
            self.lst.insert("end", describe_charge(c))

    def add_charge(self):
        d = ChargeDialog(self)
        self.win.wait_window(d)
        if d.result:
            self.charges.append(d.result)
            self.refresh_list()

    def remove_charge(self):
        s = self.lst.curselection()
        if s:
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
        self._p("What should the system produce for each client every month?", CYAN, 11)
        for text, var in (("A PDF invoice", self.o_pdf), ("A client ledger in Excel (balance + payments you enter by hand)", self.o_ledger),
                          ("A data file (JSON) for other systems", self.o_json),
                          ("A Word invoice using a starter template made for me", self.o_starter)):
            _check(self.body, text, var).pack(fill="x", padx=24, pady=2)
        _label(self.body, "...or fill in your own Word template. In it, type {{client_name}}, {{invoice_no}}, {{date}}, {{period}}, "
                          "{{total}}, {{payment_terms}} and a table row with {{item}} {{description}} {{qty}} {{unit_price}} "
                          "{{amount}}.", GREEN, 9, wraplength=800).pack(fill="x", padx=18, pady=(10, 2))
        r = tk.Frame(self.body, bg=PANEL)
        r.pack(fill="x", padx=18)
        _entry(r, self.o_template, 62).pack(side="left")
        _button(r, "BROWSE", lambda: self.o_template.set(os.path.normpath(filedialog.askopenfilename(
            filetypes=[("Word", "*.docx")]) or self.o_template.get()))).pack(side="left", padx=6)
        r = tk.Frame(self.body, bg=PANEL)
        r.pack(fill="x", padx=18, pady=14)
        _label(r, "Folder for the results", width=26).pack(side="left")
        _entry(r, self.out_dir, 46).pack(side="left")
        _button(r, "BROWSE", lambda: self.out_dir.set(os.path.normpath(filedialog.askdirectory() or self.out_dir.get()))).pack(
            side="left", padx=6)
        _label(self.body, "Each month the results go in a sub-folder named for the month. Your original files are never changed.",
               "#3a6a70", 9).pack(fill="x", padx=18)

    def step5(self):
        self._p("Review. Test it on a month before saving - nothing is written during the test.", CYAN, 11)
        r = tk.Frame(self.body, bg=PANEL)
        r.pack(fill="x", padx=18, pady=4)
        _label(r, "Month to test", width=16).pack(side="left")
        self.test_month.set(self.test_month.get() or self.period())
        _entry(r, self.test_month, 10).pack(side="left")
        _button(r, "RUN TEST", self.run_test, GREEN).pack(side="left", padx=10)
        self.out = tk.Text(self.body, font=(FONT, 9), bg=INPUT, fg=GREEN, relief="flat", wrap="none", height=22,
                           highlightthickness=1, highlightbackground="#0b5560", state="disabled")
        self.out.pack(fill="both", expand=True, padx=18, pady=6)
        for tag, col in (("bad", RED), ("good", GREEN), ("info", "#8aa7ff"), ("head", CYAN)):
            self.out.tag_configure(tag, foreground=col)
        self.say(f"Setup will be saved to: {os.path.join(self.save_dir.get(), 'config.yml')}", "info")
        self.say(f"Company: {self.company.get()}   Clients file: {os.path.basename(self.cfp.file.get())}", "info")
        for c in self.charges:
            self.say("  + " + describe_charge(c), "good")
        self.say("Press RUN TEST to read your files for the chosen month.", "info")

    def say(self, text, tag="good"):
        self.out.configure(state="normal")
        self.out.insert("end", text + "\n", tag)
        self.out.configure(state="disabled")

    # ---- answers ---------------------------------------------------------------------------------------------
    def collect(self):
        if hasattr(self, "terms") and self.terms.winfo_exists():
            self._terms = self.terms.get("1.0", "end").strip()
        if hasattr(self, "needs") and self.needs.winfo_exists():
            self._needs = self.needs.get("1.0", "end").strip()
        cl = dict(self._cl)
        cl.update({k: v for k, v in self._cm.items() if v})
        outs = {"pdf": self.o_pdf.get(), "ledger_xlsx": self.o_ledger.get(), "json": self.o_json.get()}
        if self.o_template.get().strip():
            outs["docx_template"] = self.o_template.get().strip()
        elif self.o_starter.get():
            outs["starter_template"] = True
        return {"company": self.company.get().strip(), "currency": self.currency.get().strip(),
                "payment_terms": getattr(self, "_terms", ""), "needs": getattr(self, "_needs", ""), "clients": cl,
                "charges": list(self.charges), "outputs": outs, "output_dir": self.out_dir.get().strip() or "output"}

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
            self.say("PROBLEMS FOUND - fix these (go Back) or check the month:", "bad")
            for p in problems:
                self.say("  ! " + p, "bad")
            return
        try:
            with tempfile.TemporaryDirectory() as tmp:
                results = plan(load_config(wizard.write_config(a, tmp)), period)
        except (ConfigError, BillingError, OSError) as e:
            self.say("PROBLEM: " + str(e), "bad")
            return
        ok = [r for r in results if r.status == "OK"]
        self.say(f"{len(ok)} invoice(s) ready, {sum(r.status == 'ERROR' for r in results)} need attention, "
                 f"total {sum(r.invoice.total for r in ok):,.2f}\n", "head")
        for r in results:
            line = f"{r.client_id:<10}{r.client_name[:22]:<24}{r.status:<9}" + (f"{r.invoice.number:<12}{r.invoice.total:>12,.2f}" if r.invoice else "") + f"  {r.message}"
            self.say(line, "good" if r.status == "OK" else "bad")

    def finish(self):
        a = self.collect()
        problems = wizard.check_answers(a, self.test_month.get().strip() or self.period())
        if problems and not messagebox.askyesno("BILLFLOW", "Some linked files could not be read for the test month:\n\n" +
                                                "\n".join(problems[:6]) + "\n\nSave the setup anyway?"):
            return
        try:
            path = wizard.write_config(a, self.save_dir.get().strip())
        except (ConfigError, OSError) as e:
            self.msg.configure(text=str(e), fg=RED)
            return
        self.saved_path = path
        messagebox.showinfo("BILLFLOW", f"Setup saved:\n{path}\n\nUse it from the main screen: PREVIEW, then EXECUTE.")
        if self.on_done:
            self.on_done(str(path), self.test_month.get().strip() or self.period())
        self.win.destroy()


def open_wizard(master, on_done=None):
    return Wizard(master, on_done)
