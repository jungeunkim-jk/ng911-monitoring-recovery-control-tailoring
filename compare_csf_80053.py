#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
compare_csf_80053.py  version 0.1

For each NENA-REF-012.1-2025 checklist item listed in the mapping file, the script
follows the NIST CSF 2.0 entries that NENA-REF-013.1-2025 assigns to the item into the
SP 800-53 informative references of the NIST CSF 2.0 Reference Tool, and compares the
controls reached in this way with the controls assigned by the author from the control
text ("direct reading").

Inputs (not distributed with this script):
  1. An Excel export of the CSF 2.0 Core from the NIST CSF 2.0 Reference Tool
     (csrc.nist.gov). The export carries its own "Generated Date".
  2. The mapping file mapping_input.csv.

Run:
  python compare_csf_80053.py --export csf2.xlsx --mapping mapping_input.csv --operator "Your Name"

Outputs:
  <prefix>_report.txt  run record, per-item comparison, references used, summary
  <prefix>_table.csv   per-item comparison as a table (UTF-8 with BOM, opens in Excel)

Rules applied:
  - Control identifiers are normalized: AU-02 -> AU-2, CP-09(01) -> CP-9(1).
  - A CSF entry that is a Category (for example PR.DS) is looked up on its Category row
    only. It is not expanded to the Subcategories under it.
  - A direct-reading control counts as "reached" if the same identifier is reached.
    A control enhancement whose base control is reached, but which is not itself
    listed, counts as "base reached". Anything else counts as "not reached".
  - For each control that is not reached, the script lists every CSF Subcategory in the
    export whose references include it, or reports that none does.

Requires Python 3.9 or later and openpyxl.
"""
import argparse
import csv
import datetime
import hashlib
import os
import platform
import re
import sys
import warnings

import openpyxl

VERSION = "0.1"
DEFAULT_REF = "SP 800-53 Rev 5.2.0"
OTHER_REF = "SP 800-53 Rev 5.1.1"

CTRL = re.compile(r"^([A-Z]{2})-0*(\d+)(?:\(0*(\d+)\))?$")
CSF_SUB = re.compile(r"^([A-Z]{2}\.[A-Z]{2}-\d{2})\s*:")
CSF_CAT = re.compile(r"\(([A-Z]{2}\.[A-Z]{2})\)")


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(65536), b""):
            h.update(block)
    return h.hexdigest()


def norm(c):
    c = c.strip()
    m = CTRL.match(c)
    if not m:
        return c                       # family-level entry such as "PT"
    s = "%s-%d" % (m.group(1), int(m.group(2)))
    if m.group(3):
        s += "(%d)" % int(m.group(3))
    return s


def base(c):
    return c.split("(")[0]


def sort_key(c):
    m = re.match(r"^([A-Z]{2})-(\d+)(?:\((\d+)\))?$", c)
    if not m:
        return (c, 0, 0)
    return (m.group(1), int(m.group(2)), int(m.group(3) or 0))


def refs_of(cell, label):
    out = set()
    for line in str(cell or "").splitlines():
        if line.startswith(label + ":"):
            out.add(norm(line.split(":", 1)[1]))
    return out


def read_export(path, label):
    warnings.simplefilter("ignore")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    generated = ""
    if "Introduction" in wb.sheetnames:
        for row in wb["Introduction"].iter_rows(values_only=True):
            if row and str(row[0] or "").strip() == "Generated Date":
                generated = str(row[1])
    ws = None
    for name in wb.sheetnames:
        if name.startswith("CSF"):
            ws = wb[name]
    if ws is None:
        sys.exit("No CSF sheet found in the export.")
    header_idx, cols = None, {}
    rows = list(ws.iter_rows(values_only=True))
    for i, row in enumerate(rows[:15]):
        cells = [str(c or "").strip() for c in row]
        if "Subcategory" in cells and "Informative References" in cells:
            header_idx = i
            cols = {name: cells.index(name) for name in ("Category", "Subcategory", "Informative References")}
            break
    if header_idx is None:
        sys.exit("Header row with 'Subcategory' and 'Informative References' not found.")
    sub, cat, other = {}, {}, {}
    for row in rows[header_idx + 1:]:
        ref_cell = row[cols["Informative References"]]
        c_cell = str(row[cols["Category"]] or "")
        s_cell = str(row[cols["Subcategory"]] or "")
        if c_cell.strip():
            m = CSF_CAT.search(c_cell)
            if m:
                cat[m.group(1)] = refs_of(ref_cell, label)
        m = CSF_SUB.match(s_cell.strip())
        if m:
            sub[m.group(1)] = refs_of(ref_cell, label)
            other[m.group(1)] = refs_of(ref_cell, OTHER_REF)
    return generated, sub, cat, other


def read_mapping(path):
    items = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        lines = [l for l in f if not l.lstrip().startswith("#") and l.strip()]
    for r in csv.DictReader(lines):
        csf = []
        for tok in r["csf_ref013"].split(";"):
            parts = tok.split()
            if not parts:
                continue
            strength = parts[1].lower() if len(parts) > 1 else ""
            csf.append((parts[0], strength))
        direct = [norm(x) for x in r["direct_controls"].split(";") if x.strip()]
        items.append({"item": r["item"].strip(), "section": r["sta040_section"].strip(),
                      "topic": r["topic"].strip(), "csf": csf, "direct": direct})
    return items


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--export", required=True, help="CSF 2.0 Reference Tool Excel export")
    ap.add_argument("--mapping", required=True, help="mapping_input.csv")
    ap.add_argument("--operator", required=True, help="name of the person running the script")
    ap.add_argument("--ref", default=DEFAULT_REF, help='reference label, default "%s"' % DEFAULT_REF)
    ap.add_argument("--out-prefix", default="crosswalk_v0_1")
    a = ap.parse_args()

    generated, sub, cat, other = read_export(a.export, a.ref)
    if not any(sub.values()):
        sys.exit('No references labelled "%s" found in the export.' % a.ref)
    items = read_mapping(a.mapping)

    index = {}
    for sid, refs in sub.items():
        for c in refs:
            index.setdefault(c, []).append(sid)

    results, used = [], {}
    for it in items:
        chain, notes = set(), []
        for cid, _ in it["csf"]:
            if cid in sub:
                chain |= sub[cid]
                used[cid] = sub[cid]
            elif cid in cat:
                chain |= cat[cid]
                used[cid] = cat[cid]
                notes.append("%s is a Category; its Category row lists %d %s references; Subcategories not expanded"
                             % (cid, len(cat[cid]), a.ref))
            else:
                notes.append("%s not found in the export" % cid)
        reached, base_reached, not_reached = [], [], []
        for c in it["direct"]:
            if c in chain:
                reached.append(c)
            elif "(" in c and base(c) in chain:
                base_reached.append(c)
            else:
                not_reached.append(c)
        fams = sorted({c.split("-")[0] for c in it["direct"]})
        chain_in_fams = sorted((c for c in chain if c.split("-")[0] in fams), key=sort_key)
        policy_only = bool(chain_in_fams) and all(re.match(r"^[A-Z]{2}-1$", c) for c in chain_in_fams)
        elsewhere = {c: sorted(index.get(c, [])) for c in not_reached}
        results.append(dict(it, chain=sorted(chain, key=sort_key), reached=reached, base_reached=base_reached,
                            not_reached=not_reached, chain_in_fams=chain_in_fams, policy_only=policy_only,
                            elsewhere=elsewhere, notes=notes))

    now = datetime.datetime.now().astimezone().isoformat(timespec="seconds")
    script = os.path.abspath(__file__)
    L = []
    L.append("crosswalk comparison report")
    L.append("Script: compare_csf_80053.py version %s, SHA-256 %s" % (VERSION, sha256_file(script)))
    L.append("Run: %s" % now)
    L.append("Operator: %s" % a.operator)
    L.append("Environment: Python %s, openpyxl %s, %s %s" % (platform.python_version(), openpyxl.__version__,
                                                           platform.system(), platform.release()))
    L.append("Export file: %s, SHA-256 %s" % (os.path.basename(a.export), sha256_file(a.export)))
    L.append("Export Generated Date: %s" % (generated or "not found"))
    L.append("Reference label used: %s" % a.ref)
    L.append("Mapping file: %s, SHA-256 %s" % (os.path.basename(a.mapping), sha256_file(a.mapping)))
    L.append("")
    L.append("Per-item comparison")
    L.append("=" * 78)
    for r in results:
        L.append("Item %s  (STA-040.2 section %s)  %s" % (r["item"], r["section"], r["topic"]))
        L.append("  CSF entries (REF-013.1): %s" % "; ".join("%s %s" % x for x in r["csf"]))
        L.append("  Controls reached through the CSF references (%d): %s" % (len(r["chain"]), ", ".join(r["chain"]) or "none"))
        L.append("  Direct reading: %s" % ", ".join(r["direct"]))
        L.append("    reached: %s" % (", ".join(r["reached"]) or "none"))
        L.append("    base control reached, enhancement not listed: %s" % (", ".join(r["base_reached"]) or "none"))
        L.append("    not reached: %s" % (", ".join(r["not_reached"]) or "none"))
        for c, subs in r["elsewhere"].items():
            L.append("      %s is referenced by: %s" % (c, ", ".join(subs) if subs else "no Subcategory in the export"))
        if r["policy_only"]:
            L.append("  In the families of the direct reading (%s), the CSF references reach only: %s"
                     % (", ".join(sorted({c.split('-')[0] for c in r['direct']})), ", ".join(r["chain_in_fams"])))
        for n in r["notes"]:
            L.append("  Note: %s" % n)
        L.append("")

    L.append("References used (%s), by CSF entry" % a.ref)
    L.append("=" * 78)
    for cid in sorted(used):
        line = "%s: %s" % (cid, ", ".join(sorted(used[cid], key=sort_key)) or "none")
        if cid in other and other[cid] != used[cid]:
            line += "   [%s differs: only in %s: %s; only in %s: %s]" % (
                OTHER_REF, a.ref, ", ".join(sorted(used[cid] - other[cid], key=sort_key)) or "none",
                OTHER_REF, ", ".join(sorted(other[cid] - used[cid], key=sort_key)) or "none")
        L.append(line)
    L.append("")

    n = len(results)
    any_hit = [r for r in results if r["reached"] or r["base_reached"]]
    none_hit = [r for r in results if not (r["reached"] or r["base_reached"])]
    pol = [r for r in none_hit if r["policy_only"]]
    total = sum(len(r["direct"]) for r in results)
    n_reached = sum(len(r["reached"]) for r in results)
    n_base = sum(len(r["base_reached"]) for r in results)
    n_not = sum(len(r["not_reached"]) for r in results)
    nowhere = sorted({c for r in results for c, s in r["elsewhere"].items() if not s}, key=sort_key)
    nowhere_base = [c for c in nowhere if "(" not in c]
    nowhere_enh = [c for c in nowhere if "(" in c]
    all_enh = sorted({c for c in index if "(" in c}, key=sort_key)
    L.append("Summary")
    L.append("=" * 78)
    L.append("Items: %d" % n)
    L.append("Items with at least one direct-reading control reached (or its base control): %d (%s)"
             % (len(any_hit), ", ".join(r["item"] for r in any_hit) or "none"))
    L.append("Items with no direct-reading control reached: %d (%s)"
             % (len(none_hit), ", ".join(r["item"] for r in none_hit) or "none"))
    L.append("  of which the CSF references reach only the -1 control in the families of the direct reading: %d (%s)"
             % (len(pol), ", ".join(r["item"] for r in pol) or "none"))
    L.append("Direct-reading assignments: %d; reached %d; base control reached %d; not reached %d"
             % (total, n_reached, n_base, n_not))
    L.append("Direct-reading base controls referenced by no Subcategory in the export: %s" % (", ".join(nowhere_base) or "none"))
    L.append("Direct-reading control enhancements referenced by no Subcategory in the export: %s" % (", ".join(nowhere_enh) or "none"))
    L.append("  (the export lists %d control enhancements in total under %s: %s)" % (len(all_enh), a.ref, ", ".join(all_enh) or "none"))

    rep = a.out_prefix + "_report.txt"
    with open(rep, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(L) + "\n")
    tab = a.out_prefix + "_table.csv"
    with open(tab, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Item", "STA-040.2 section", "Topic", "CSF entries (REF-013.1)",
                    "Controls reached through CSF references", "Direct reading", "Reached",
                    "Base control reached", "Not reached", "Not-reached controls referenced by"])
        for r in results:
            w.writerow([r["item"], r["section"], r["topic"], "; ".join("%s %s" % x for x in r["csf"]),
                        ", ".join(r["chain"]), ", ".join(r["direct"]), ", ".join(r["reached"]),
                        ", ".join(r["base_reached"]), ", ".join(r["not_reached"]),
                        "; ".join("%s: %s" % (c, ", ".join(s) if s else "none") for c, s in r["elsewhere"].items())])
    print("Wrote %s and %s" % (rep, tab))


if __name__ == "__main__":
    main()
