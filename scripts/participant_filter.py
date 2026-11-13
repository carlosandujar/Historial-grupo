"""Normaliza apellidos y compara los formatos de nombres del historial."""

import re
from unicodedata import category, normalize


DASH = re.compile(r"[-‐‑‒–—−]+")
AFFILIATION = re.compile(r"\([^()]*\)")
TRAILING_INITIALS = re.compile(r"(?:\s+[^\W\d_]\.)+$")


def surname_keys(value):
    folded = normalize("NFKD", value.casefold())
    plain = "".join(char for char in folded if not category(char).startswith("M"))
    parts = [" ".join(re.findall(r"[^\W\d_]+", part)) for part in DASH.split(plain)]
    keys = {part for part in parts if part}
    if len(parts) > 1 and keys:
        keys.add(" ".join(part for part in parts if part))
    return keys


def participant_names(value):
    """Lee los apellidos del filtro separados por comas."""
    return {key for surname in (value or "").split(",") for key in surname_keys(surname)}


def people(value, *, ips):
    for group in re.split(r"[;\r\n]+", str(value or "")):
        if not ips:
            yield group
            continue
        # En IPs, una afiliación seguida de coma marca el final de una persona.
        entries = re.split(r"(?<=\))\s*,\s*", group)
        if len(entries) == 1 and not AFFILIATION.search(group):
            pieces = group.split(",")
            if len(pieces) > 1 and all(len(AFFILIATION.sub("", part).split()) > 1 for part in pieces):
                entries = pieces
        yield from entries


def matches_participants(value, selected, *, ips=False):
    if not selected:
        return True
    for person in people(value, ips=ips):
        name = AFFILIATION.sub("", person).strip()
        family, comma, _ = name.partition(",")
        if not comma:
            # Para nombres sin coma, el último término es el apellido.
            # Una inicial al final (p. ej. Maravita A.) no forma parte de él.
            terms = TRAILING_INITIALS.sub("", name).split()
            family = terms[-1] if terms else ""
        if selected & surname_keys(family):
            return True
    return False
