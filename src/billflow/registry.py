"""Plugin registries. Companies extend billflow by registering their own pieces."""

SOURCES = {}   # name -> fn(params, ctx) -> list[dict]      (where data comes from)
RULES = {}     # name -> fn(client, params, ctx) -> list[Line]  (how a charge is calculated)
OUTPUTS = {}   # name -> fn(invoice, params, ctx) -> list[path]  (what gets produced)


def _register(table, name):
    def deco(fn):
        table[name] = fn
        return fn
    return deco


def source(name):
    return _register(SOURCES, name)


def rule(name):
    return _register(RULES, name)


def output(name):
    return _register(OUTPUTS, name)
