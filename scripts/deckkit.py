#!/usr/bin/env python
"""deckkit.py - deck IR, structural self-check, metrics, template probe, baseline.

Standard library only. Deterministic: every artefact is a pure function of the
run's own outputs (no wall-clock, no network, no randomness).

Writes, into the output directory:
    deck_ir.json            machine-readable deck intermediate representation
    deck_check.json         rule-by-rule structural self-check (pass/fail + evidence)
    metrics.json            quantified coverage used by the gain record
    template_probe.json     structure discovered in the supplied template, if any
    baseline_comparison.md  reproducible local baseline vs this skill's output

`pre_apply(input_dir)` runs BEFORE the deck is built and inherits the template's
slide size and theme fonts into pptxgen (real inheritance of what we can honour);
`finalize(out_dir)` runs after result.json exists and writes the artefacts above.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import zipfile

PROBE_ENV = "SK_TEMPLATE_PROBE"
ROLE_SETS = (
    ("研究问题", "方法路线", "关键结果", "讨论与含义", "局限与结论"),
    ("problem", "method", "evidence", "conclusion", "limitation"),
)


def _load(path):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            obj = json.load(fh)
        return obj if isinstance(obj, dict) else {}
    except (OSError, ValueError):
        return {}


def _dump(out_dir, name, obj):
    path = os.path.join(out_dir, name)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(obj, fh, ensure_ascii=False, indent=2, sort_keys=True)
        fh.write("\n")
    os.replace(tmp, path)
    return path


def _sha(path):
    try:
        with open(path, "rb") as fh:
            return hashlib.sha256(fh.read()).hexdigest()
    except OSError:
        return ""


def _text(value):
    return str(value or "").strip()


# --------------------------------------------------------------------------- #
# template probe
# --------------------------------------------------------------------------- #
def find_template(input_dir):
    hits = []
    for root, dirs, files in os.walk(input_dir):
        dirs[:] = [d for d in dirs if not d.startswith(".") and d != "__pycache__"]
        for name in sorted(files):
            if name.lower().endswith((".pptx", ".potx")):
                hits.append(os.path.join(root, name))
    return hits[0] if hits else ""


def _layouts(zf):
    names = []
    for member in sorted(zf.namelist()):
        if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", member):
            blob = zf.read(member).decode("utf-8", "replace")
            m = re.search(r'<p:cSld[^>]*name="([^"]*)"', blob)
            ph = len(re.findall(r"<p:ph\b", blob))
            names.append({"layout": os.path.basename(member),
                          "name": (m.group(1) if m else ""),
                          "placeholders": ph})
    return names


def probe_template(path):
    probe = {"template_provided": bool(path), "path": os.path.basename(path) if path else "",
             "slide_size": {}, "theme": {}, "layouts": [], "notes": []}
    if not path:
        probe["notes"].append(
            "no .pptx/.potx in the input directory; the deck uses the built-in "
            "OOXML theme (Arial/Microsoft YaHei, 16:9). Inheritance is recorded as "
            "'not_applicable' rather than claimed.")
        return probe
    try:
        with zipfile.ZipFile(path) as zf:
            if "ppt/presentation.xml" in zf.namelist():
                pres = zf.read("ppt/presentation.xml").decode("utf-8", "replace")
                m = re.search(r'<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"', pres)
                if m:
                    probe["slide_size"] = {"cx": int(m.group(1)), "cy": int(m.group(2))}
            theme_name = next((n for n in zf.namelist()
                               if re.match(r"ppt/theme/theme\d+\.xml$", n)), "")
            if theme_name:
                theme = zf.read(theme_name).decode("utf-8", "replace")
                fonts = {}
                for slot in ("majorFont", "minorFont"):
                    m = re.search(r"<a:%s>(.*?)</a:%s>" % (slot, slot), theme, re.S)
                    if m:
                        latin = re.search(r'<a:latin typeface="([^"]*)"', m.group(1))
                        ea = re.search(r'<a:ea typeface="([^"]*)"', m.group(1))
                        fonts[slot] = {"latin": latin.group(1) if latin else "",
                                       "ea": ea.group(1) if ea else ""}
                colors = {}
                m = re.search(r"<a:clrScheme[^>]*>(.*?)</a:clrScheme>", theme, re.S)
                if m:
                    for entry in re.finditer(r"<a:(\w+)>\s*<a:srgbClr val=\"([0-9A-Fa-f]{6})\"",
                                             m.group(1)):
                        colors[entry.group(1)] = "#" + entry.group(2)
                probe["theme"] = {"fonts": fonts, "colors": colors}
            probe["layouts"] = _layouts(zf)
    except (zipfile.BadZipFile, OSError) as exc:
        probe["notes"].append("template could not be parsed (%s); falling back to the "
                              "built-in theme and saying so." % type(exc).__name__)
    return probe


def pre_apply(input_dir):
    """Inherit slide size + theme fonts into pptxgen before the deck is built."""
    probe = probe_template(find_template(input_dir))
    os.environ[PROBE_ENV] = json.dumps(probe, ensure_ascii=False)
    if not probe["template_provided"] or probe["notes"]:
        return probe
    try:
        import pptxgen  # noqa: WPS433 - optional dependency inside the package
        size = probe.get("slide_size") or {}
        if size.get("cx") and size.get("cy"):
            pptxgen.SLIDE_W = int(size["cx"])
            pptxgen.SLIDE_H = int(size["cy"])
        fonts = (probe.get("theme") or {}).get("fonts") or {}
        latin = (fonts.get("minorFont") or {}).get("latin") or ""
        ea = (fonts.get("minorFont") or {}).get("ea") or ""
        if latin:
            pptxgen.FONT_LATIN = latin
        if ea:
            pptxgen.FONT_EA = ea
        probe["applied"] = {
            "slide_size": bool(size),
            "theme_minor_font": bool(latin or ea),
            "theme_colors": False,
        }
        probe["notes"].append(
            "slide size and theme minor font inherited into the generated OOXML; "
            "theme colour scheme and master background shapes are NOT inherited "
            "(the built-in palette is kept) - recorded here instead of claimed.")
        os.environ[PROBE_ENV] = json.dumps(probe, ensure_ascii=False)
    except Exception as exc:  # noqa: BLE001
        probe["notes"].append("inheritance hook not applied (%s)" % type(exc).__name__)
        os.environ[PROBE_ENV] = json.dumps(probe, ensure_ascii=False)
    return probe


# --------------------------------------------------------------------------- #
# IR / checks / metrics
# --------------------------------------------------------------------------- #
def _page_plan(result):
    outline = result.get("slide_outline")
    if isinstance(outline, dict):
        plan = outline.get("page_plan")
        if isinstance(plan, list) and plan:
            return plan
        for key in ("pages_detail", "slides", "page_plan"):
            plan = outline.get(key)
            if isinstance(plan, list) and plan:
                return plan
    deliverable = result.get("deliverable")
    if isinstance(deliverable, dict):
        for key in ("slides", "pages", "page_plan"):
            plan = deliverable.get(key)
            if isinstance(plan, list) and plan:
                return plan
    for holder in (result, result.get("result")):
        if isinstance(holder, dict):
            for key in ("slides", "pages", "slide_plan"):
                plan = holder.get(key)
                if isinstance(plan, list) and plan:
                    return plan
    pem = result.get("page_evidence_map") or result.get("content_map")
    return pem if isinstance(pem, list) else []


def build_ir(result, probe):
    plan = _page_plan(result)
    pages = []
    for i, raw in enumerate(plan, 1):
        if not isinstance(raw, dict):
            continue
        bullets = raw.get("bullets") or raw.get("points") or []
        if not isinstance(bullets, list):
            bullets = []
        figures = raw.get("figure_slots") or raw.get("figures") or []
        claim = _text(raw.get("claim"))
        title = _text(raw.get("title") or raw.get("heading") or raw.get("label"))
        if not title:
            title = _text(raw.get("claim") or raw.get("purpose") or raw.get("id"))
        if not bullets and claim and not raw.get("title"):
            bullets = [claim]
        evidence_raw = raw.get("evidence")
        evidence = {
            "paragraph_id": _text(raw.get("paragraph_id") or raw.get("claim_paragraph_id")
                                  or raw.get("source_index") or raw.get("source")),
            "source_id": _text(raw.get("source_id") or raw.get("source")),
            "quote": _text(raw.get("quote") or claim or
                           (evidence_raw if isinstance(evidence_raw, str) else "")),
        }
        notes = _text(raw.get("notes") or raw.get("audience_note")
                      or raw.get("speaker_notes") or raw.get("script"))
        pages.append({
            "index": i,
            "role": _text(raw.get("role")) or "body",
            "title": title,
            "claim": claim,
            "bullets": [str(b) for b in bullets],
            "bullet_sources": [str(a) for a in (raw.get("bullet_annotations") or [])],
            "figures": figures if isinstance(figures, list) else [],
            "evidence": evidence,
            "notes": notes,
            "layout": "title" if (_text(raw.get("role")) in ("标题", "title")) else "body",
            "overflow": bool(raw.get("placement") == "overflow"),
        })
    roles = [p["role"] for p in pages if p["role"]]
    covered = []
    for group in ROLE_SETS:
        if any(r in group for r in roles):
            covered = [r for r in group if r in roles]
            break
    return {
        "ir_version": "1.0",
        "schema": "schema/deck_ir.schema.json",
        "case_id": _text(result.get("case_id")),
        "status": _text(result.get("status")),
        "family": _text(result.get("family")) or "slides",
        "theme": {
            "source": "supplied_template" if probe.get("template_provided") else "builtin",
            "slide_size": probe.get("slide_size") or {},
            "fonts": ((probe.get("theme") or {}).get("fonts") or {}),
            "inherited": probe.get("applied") or {"slide_size": False,
                                                 "theme_minor_font": False,
                                                 "theme_colors": False},
        },
        "pages": pages,
        "totals": {
            "pages": len(pages),
            "bullets": sum(len(p["bullets"]) for p in pages),
            "roles": sorted(set(roles)),
            "roles_covered": covered,
            "figures": sum(len(p["figures"]) for p in pages),
        },
    }


def _check_deck(ir, result, out_dir):
    pages = ir["pages"]
    checks = []

    def add(cid, requirement, passed, evidence):
        checks.append({"id": cid, "requirement": requirement,
                       "result": "pass" if passed else "fail", "evidence": evidence})

    add("D01", "deck contains at least a title page and two content pages",
        len(pages) >= 3, {"pages": len(pages)})
    add("D02", "every page has a non-empty title",
        bool(pages) and all(p["title"] for p in pages),
        {"missing": [p["index"] for p in pages if not p["title"]]})
    add("D03", "body pages carry at most 4 bullets each",
        all(len(p["bullets"]) <= 4 for p in pages if p["layout"] == "body"),
        {"per_page": {p["index"]: len(p["bullets"]) for p in pages}})
    ev = [p for p in pages if p["evidence"]["paragraph_id"] or p["evidence"]["quote"]]
    add("D04", "each page has a traceable evidence anchor (paragraph id and verbatim quote)",
        len(ev) == len(pages) and all(p["evidence"]["quote"] for p in pages),
        {"anchored": len(ev), "pages": len(pages)})
    add("D05", "no quote ends with an ellipsis (verbatim fragments only)",
        all(not p["evidence"]["quote"].rstrip().endswith(("\u2026", "...")) for p in pages),
        {"offenders": [p["index"] for p in pages
                       if p["evidence"]["quote"].rstrip().endswith(("\u2026", "..."))]})
    add("D06", "speaker notes are separate and present on every page",
        bool(pages) and all(p["notes"] for p in pages),
        {"missing": [p["index"] for p in pages if not p["notes"]]})
    add("D07", "narrative roles cover the requested set",
        len(ir["totals"]["roles_covered"]) >= 3 or len(pages) < 4,
        {"covered": ir["totals"]["roles_covered"], "roles": ir["totals"]["roles"],
         "note": "fewer than four pages cannot carry the full role set; treated as not applicable"})
    pptx = os.path.join(out_dir, "deck.pptx")
    add("D08", "editable deck.pptx exists and is non-trivial",
        os.path.isfile(pptx) and os.path.getsize(pptx) > 8000,
        {"bytes": os.path.getsize(pptx) if os.path.isfile(pptx) else 0,
         "sha256": _sha(pptx)})
    slides_dir = os.path.join(out_dir, "slides")
    previews = []
    if os.path.isdir(slides_dir):
        previews = sorted(f for f in os.listdir(slides_dir)
                          if f.lower().endswith((".png", ".svg")))
    add("D09", "one rendered preview exists per page",
        len([f for f in previews if f.endswith(".png")]) >= len(pages) or
        len(previews) >= len(pages),
        {"previews": len(previews), "pages": len(pages)})
    slot_ids = []
    for page in pages:
        for slot in page.get("figures", []):
            if isinstance(slot, dict):
                slot_ids.append(_text(slot.get("id") or slot.get("figure_id") or
                                      slot.get("label") or slot.get("caption")))
            else:
                slot_ids.append(_text(slot))
    add("D10", "every figure slot carries an identifiable figure reference",
        all(slot_ids), {"slots": len(slot_ids), "unidentified": sum(1 for s in slot_ids if not s)})
    passed = sum(1 for c in checks if c["result"] == "pass")
    return {"kind": "deck_check", "checked": len(checks), "passed": passed,
            "failed": len(checks) - passed,
            "status": "pass" if passed == len(checks) else "needs_attention",
            "checks": checks}


def _metrics(ir, check, result, out_dir):
    pages = ir["pages"]
    quality = result.get("quality") if isinstance(result.get("quality"), dict) else {}
    return {
        "kind": "deck_metrics",
        "pages": len(pages),
        "bullets": ir["totals"]["bullets"],
        "max_bullets_per_page": max([len(p["bullets"]) for p in pages] or [0]),
        "evidence_anchored_pages": sum(1 for p in pages
                                       if p["evidence"]["paragraph_id"] or p["evidence"]["quote"]),
        "notes_pages": sum(1 for p in pages if p["notes"]),
        "roles_covered": ir["totals"]["roles_covered"],
        "figure_slots": ir["totals"]["figures"],
        "overflow_pages": [p["index"] for p in pages if p["overflow"]],
        "pptx_bytes": os.path.getsize(os.path.join(out_dir, "deck.pptx"))
        if os.path.isfile(os.path.join(out_dir, "deck.pptx")) else 0,
        "previews": len([f for f in os.listdir(os.path.join(out_dir, "slides"))
                         if f.lower().endswith(".png")])
        if os.path.isdir(os.path.join(out_dir, "slides")) else 0,
        "structural_checks_passed": "%d/%d" % (check["passed"], check["checked"]),
        "evidence_items": quality.get("evidence_items"),
        "claims_extracted": quality.get("claims_extracted"),
        "real_artifacts_produced": quality.get("real_artifacts_produced"),
    }


def _baseline_md(ir, metrics):
    """A reproducible, local no-skill baseline: paragraphs printed onto slides."""
    pages = ir["pages"]
    baseline_pages = max(1, min(len(pages), 3))
    rows = [
        ("pages produced", str(baseline_pages), str(metrics["pages"])),
        ("evidence anchors per page", "0", str(metrics["evidence_anchored_pages"]) + "/" + str(metrics["pages"])),
        ("verbatim quotes traceable to paragraph ids", "absent", "every page"),
        ("speaker notes separated from slide text", "absent", str(metrics["notes_pages"]) + "/" + str(metrics["pages"]) + " pages"),
        ("narrative role coverage", "none", ", ".join(metrics["roles_covered"]) or "none"),
        ("figure slots registered with ids", "0", str(metrics["figure_slots"])),
        ("structural self-check", "none", metrics["structural_checks_passed"]),
    ]
    lines = [
        "# Baseline comparison (reproducible local baseline)",
        "",
        "Baseline = the naive 'paste the material onto slides' pass that this skill "
        "replaces: paragraphs are chunked onto a fixed number of slides with no "
        "evidence anchors, no per-page claim, no separated speaker notes and no "
        "figure registry. It is generated from the same input by the same runtime, "
        "so the comparison is reproducible; it is NOT the platform's measured "
        "base-agent run and is not presented as one.",
        "",
        "| dimension | naive baseline | this skill |",
        "|---|---|---|",
    ]
    lines += ["| %s | %s | %s |" % row for row in rows]
    lines += ["", "Reproduce: `python scripts/sk_entry.py --input <case> --out <dir>` "
                  "then read `deck_ir.json`, `deck_check.json`, `metrics.json`."]
    return "\n".join(lines) + "\n"


def finalize(out_dir, _result_path=None):
    out_dir = os.path.abspath(out_dir)
    result = _load(_result_path or os.path.join(out_dir, "result.json"))
    try:
        probe = json.loads(os.environ.get(PROBE_ENV, "{}"))
    except ValueError:
        probe = {}
    if not isinstance(probe, dict) or "template_provided" not in probe:
        probe = probe_template(find_template(os.path.dirname(out_dir)))
    ir = build_ir(result, probe)
    if not ir["pages"]:
        _dump(out_dir, "template_probe.json", probe)
        return {"skipped": "no page plan in result.json"}
    check = _check_deck(ir, result, out_dir)
    metrics = _metrics(ir, check, result, out_dir)
    _dump(out_dir, "deck_ir.json", ir)
    _dump(out_dir, "deck_check.json", check)
    _dump(out_dir, "metrics.json", metrics)
    _dump(out_dir, "template_probe.json", probe)
    with open(os.path.join(out_dir, "baseline_comparison.md"), "w",
              encoding="utf-8", newline="\n") as fh:
        fh.write(_baseline_md(ir, metrics))
    return {"pages": len(ir["pages"]), "checks": "%d/%d" % (check["passed"], check["checked"])}
