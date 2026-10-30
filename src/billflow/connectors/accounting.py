"""Connectors that hand invoices to other systems: Xero, QuickBooks Online and a generic webhook.

Safety: the request body is always saved as a JSON file first. Nothing is sent unless the output has `mode: api`
AND the run was started with --push. API tokens are read from environment variables, never from the config.
"""
import json
import os
import urllib.error
import urllib.request
from datetime import timedelta

from ..models import BillingError
from ..registry import output
from .outputs import _fmt_path, _money


def _post_json(url, headers, payload):
    req = urllib.request.Request(url, data=json.dumps(payload).encode(), method="POST",
                                 headers={"Content-Type": "application/json", "Accept": "application/json",
                                          "User-Agent": "billflow", **headers})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = resp.read().decode("utf-8", "replace")
            return json.loads(body) if body else {}
    except urllib.error.HTTPError as e:
        raise BillingError(f"{url.split('/')[2]} rejected the invoice ({e.code}): {e.read().decode('utf-8', 'replace')[:300]}")
    except (OSError, ValueError) as e:
        raise BillingError(f"Could not reach {url.split('/')[2]}: {e}")


def _env(params, key, default):
    name = params.get(key, default)
    val = os.environ.get(name, "")
    if not val:
        raise BillingError(f"Environment variable {name} is not set")
    return val


def invoice_dict(invoice, ctx):
    return {"company": ctx.cfg["company"]["name"], "currency": ctx.cfg["company"].get("currency", ""),
            "invoice_no": invoice.number, "client_id": invoice.client_id, "client_name": invoice.client_name,
            "issue_date": invoice.issue_date.isoformat(), "period": invoice.period,
            "lines": [{"item": l.item, "description": l.description, "qty": str(l.qty),
                       "unit_price": str(l.unit_price), "amount": str(l.amount)} for l in invoice.lines],
            "total": str(invoice.total)}


def _deliver(invoice, params, ctx, name, build):
    """build() -> (url, headers, payload). Saves the payload; sends only in api mode with --push."""
    url, headers, payload = build()
    path = _fmt_path(ctx, params, "draft_path", invoice, "{period}/{client_id}/" + name + "-{client_id}.json")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    if params.get("mode", "draft") == "api" and ctx.cfg.get("_push"):
        _post_json(url, headers, payload)
    return [path]


def _due(invoice, params):
    return (invoice.issue_date + timedelta(days=int(params.get("due_days", 7)))).isoformat()


@output("xero")
def xero_output(invoice, params, ctx):
    """Creates the invoice in Xero (Accounts Receivable, DRAFT by default).
    Env: XERO_ACCESS_TOKEN, XERO_TENANT_ID (names can be changed with token_env / tenant_env)."""
    def build():
        contact = {"ContactID": invoice.client[params["contact_id_column"]]} if params.get("contact_id_column") \
            else {"Name": invoice.client_name}
        payload = {"Invoices": [{
            "Type": "ACCREC", "Contact": contact, "Date": invoice.issue_date.isoformat(), "DueDate": _due(invoice, params),
            "InvoiceNumber": invoice.number, "Reference": invoice.period, "Status": params.get("status", "DRAFT"),
            "LineAmountTypes": "NoTax", "CurrencyCode": ctx.cfg["company"].get("currency") or None,
            "LineItems": [{"Description": l.description or l.item, "Quantity": float(l.qty), "UnitAmount": float(l.unit_price),
                           **({"AccountCode": params["account_code"]} if params.get("account_code") else {})}
                          for l in invoice.lines]}]}
        if not payload["Invoices"][0]["CurrencyCode"]:
            del payload["Invoices"][0]["CurrencyCode"]
        headers = {}
        if params.get("mode") == "api" and ctx.cfg.get("_push"):
            headers = {"Authorization": "Bearer " + _env(params, "token_env", "XERO_ACCESS_TOKEN"),
                       "xero-tenant-id": _env(params, "tenant_env", "XERO_TENANT_ID")}
        return "https://api.xero.com/api.xro/2.0/Invoices", headers, payload
    return _deliver(invoice, params, ctx, "xero", build)


@output("quickbooks")
def quickbooks_output(invoice, params, ctx):
    """Creates the invoice in QuickBooks Online. Needs `customer_id_column` (QuickBooks customer id in your client list)
    and `item_id` (the QuickBooks product/service to book charges against).
    Env: QUICKBOOKS_ACCESS_TOKEN, QUICKBOOKS_REALM_ID. Use `base_url` for the sandbox."""
    def build():
        cust = str(invoice.client.get(params.get("customer_id_column", ""), "")).strip()
        if not cust:
            raise BillingError(f"No QuickBooks customer id for client {invoice.client_id}")
        payload = {"DocNumber": invoice.number[:21], "TxnDate": invoice.issue_date.isoformat(),
                   "DueDate": _due(invoice, params), "CustomerRef": {"value": cust}, "PrivateNote": invoice.period,
                   "Line": [{"DetailType": "SalesItemLineDetail", "Amount": float(l.amount),
                             "Description": l.description or l.item,
                             "SalesItemLineDetail": {"ItemRef": {"value": str(params["item_id"])},
                                                     "Qty": float(l.qty), "UnitPrice": float(l.unit_price)}}
                            for l in invoice.lines]}
        headers, realm = {}, "REALM"
        if params.get("mode") == "api" and ctx.cfg.get("_push"):
            headers = {"Authorization": "Bearer " + _env(params, "token_env", "QUICKBOOKS_ACCESS_TOKEN")}
            realm = _env(params, "realm_env", "QUICKBOOKS_REALM_ID")
        base = params.get("base_url", "https://quickbooks.api.intuit.com").rstrip("/")
        return f"{base}/v3/company/{realm}/invoice?minorversion=65", headers, payload
    return _deliver(invoice, params, ctx, "quickbooks", build)


@output("webhook")
def webhook_output(invoice, params, ctx):
    """POSTs the invoice as JSON to any URL (Zapier, Make, n8n, your own system). Optional `bearer_env` for a token."""
    def build():
        headers = dict(params.get("headers", {}))
        if params.get("mode") == "api" and ctx.cfg.get("_push") and params.get("bearer_env"):
            headers["Authorization"] = "Bearer " + _env(params, "bearer_env", "")
        return params["url"], headers, invoice_dict(invoice, ctx)
    return _deliver(invoice, params, ctx, "webhook", build)
