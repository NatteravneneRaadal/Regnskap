"""Lager regnskapssiden (index.html) fra regnskapsarbeidsboka.

Bruk:  python generate_regnskap.py <regnskap.xlsx eller mappe> <mal.html> <ut.html> [år]

Leser arkene "<år> Hovedbok" og "<år> Årsregnskap". Uten år brukes det
høyeste året som har begge arkene. Er første argument en mappe, brukes
Excel-fila i mappa som har det nyeste året.
"""
import glob
import json
import os
import re
import sys
from datetime import datetime

import openpyxl

SEKSJONER = {"Driftsinntekter", "Driftskostnader", "Finansposter",
             "Eiendeler", "Egenkapital og gjeld"}


def finn_aar(wb):
    aar = set()
    for n in wb.sheetnames:
        m = re.match(r"(\d{4}) Hovedbok$", n)
        if m and f"{m.group(1)} Årsregnskap" in wb.sheetnames:
            aar.add(int(m.group(1)))
    return max(aar, default=0)


def les_hovedbok(ws):
    rader = list(ws.iter_rows(min_row=6, values_only=True))
    ib, tx = {}, []
    i = 0
    while i < len(rader) - 1:
        r = rader[i]
        if r[1] is None or r[2] is None:
            i += 1
            continue
        mot = rader[i + 1]
        inn, ut = r[4] or 0, r[5] or 0
        if r[2] == "IB":
            ib[mot[2]] = round(ut or inn, 2)
        else:
            tx.append({"b": r[0], "d": r[1].strftime("%Y-%m-%d"), "k": r[2],
                       "m": mot[2], "t": r[3] or "",
                       "inn": round(inn, 2), "ut": round(ut, 2)})
        i += 2
    return ib, tx


def les_aarsregnskap(ws):
    oppstillinger, cur, forventer_aar = [], None, False
    for r in ws.iter_rows(values_only=True):
        tekst, a, b = r[1], r[2], r[3]
        if forventer_aar and a is not None:
            cur["kolonner"] = [a, b]
            forventer_aar = False
            continue
        if tekst is None:
            continue
        tekst = str(tekst).strip()
        if " - " in tekst and a is None:
            cur = {"tittel": tekst.split(" - ")[0], "kolonner": [], "rader": []}
            oppstillinger.append(cur)
            forventer_aar = True
        elif tekst in SEKSJONER:
            cur["rader"].append({"type": "seksjon", "tekst": tekst})
        else:
            typ = "sum" if tekst.startswith("Sum") or tekst == "Årsresultat" else "linje"
            cur["rader"].append({"type": typ, "tekst": tekst,
                                 "v": [round(a or 0, 2), round(b or 0, 2)]})
    return oppstillinger


def velg_fil(kilde):
    if not os.path.isdir(kilde):
        return kilde
    filer = sorted(glob.glob(os.path.join(kilde, "*.xlsx")))
    filer = [f for f in filer if not os.path.basename(f).startswith("~$")]
    if not filer:
        sys.exit(f"Fant ingen .xlsx-fil i {kilde}")
    beste = max(filer, key=lambda f: finn_aar(openpyxl.load_workbook(f, read_only=True)))
    print(f"Bruker {beste}")
    return beste


def main():
    kilde, mal, ut = sys.argv[1:4]
    wb = openpyxl.load_workbook(velg_fil(kilde), data_only=True)
    aar = int(sys.argv[4]) if len(sys.argv) > 4 else finn_aar(wb)
    ib, tx = les_hovedbok(wb[f"{aar} Hovedbok"])
    oppst = les_aarsregnskap(wb[f"{aar} Årsregnskap"])
    data = {"aar": aar, "ib": ib, "tx": tx, "oppstillinger": oppst,
            "sisteBilag": max(t["d"] for t in tx) if tx else None,
            "generert": datetime.now().strftime("%Y-%m-%d")}
    html = open(mal, encoding="utf-8").read()
    # Én transaksjon per linje, så ingen linje blir for lang
    js = json.dumps(data, ensure_ascii=False).replace('{"b":', '\n{"b":')
    html = html.replace("/*__DATA__*/null", js)
    open(ut, "w", encoding="utf-8").write(html)
    print(f"Skrev {ut}: {aar}, {len(tx)} bilag")


if __name__ == "__main__":
    main()
