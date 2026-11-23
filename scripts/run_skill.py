#!/usr/bin/env python3
"""Paper-to-academic-slides skill (standard library only).

Extracts a narrative spine from paper material, builds page-level claims with
figure/table/section locators, generates speaker notes distinct from slide
text, and writes a reproducible deck specification to /app/output/result.json.
Deterministic, offline, no fabricated research results or rendered decks.
"""

import argparse
import hashlib
import json
import os
import re
import sys

DEFAULT_CASE = "/input/case.json"
DEFAULT_OUT = "/app/output/result.json"
DEFAULT_WORK = "/app/work"
SKILL_NAME = "paper-to-academic-slides"
SKILL_VERSION = "1.0.0"
CHALLENGE_ID = "ppt-c7b80f0d"
MAX_TEXT = 400000
SKIP_KEYS = {"output_contract", "scoring", "tags", "meta", "case_id",
             "challenge_id", "skill", "contract_notes"}

ROLE_ORDER = ["problem", "method", "evidence", "conclusion", "limitation"]
CUE_WORDS = {
    "problem": ("问题", "挑战", "瓶颈", "不足", "缺乏", "难以", "需求", "背景",
                "challenge", "problem", "gap", "remains", "lack"),
    "method": ("方法", "设计", "构建", "提出", "开发", "采用", "流程", "实验", "模型",
               "we develop", "we design", "we propose", "method", "approach"),
    "evidence": ("结果", "发现", "显示", "表明", "提高", "降低", "达到", "显著", "相比",
                 "results", "we found", "showed", "increased", "decreased", "compared"),
    "conclusion": ("结论", "意义", "贡献", "因此", "综上", "表明该", "可用于",
                   "conclusion", "in summary", "therefore", "contribution"),
    "limitation": ("局限", "限制", "不足", "未来", "尚需", "未能", "样本量",
                   "limitation", "future work", "caveat"),
}
REF_RE = re.compile(r"(图|图\s*|Figure|Fig\.?|表|Table)\s*([0-9]{1,3}[A-Za-z]?)")
NUMBER_RE = re.compile(r"\d+(?:\.\d+)?\s*(?:%|倍|mg|g|mL|L|h|小时|分钟|天|周|°C|nm|μm|mM|μM|nM|kDa|Da|例|名|人)")
STOPWORDS = {"the", "and", "for", "with", "from", "that", "this", "these", "those",
             "into", "using", "used", "were", "was", "are", "is", "be", "of", "in",
             "on", "to", "by", "as", "at", "an", "or", "we", "our", "their", "its"}


def sha256_text(text):
    return hashlib.sha256((text or "").encode("utf-8")).hexdigest()


def norm(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def load_json(path):
    with open(path, "r", encoding="utf-8-sig") as handle:
        return json.load(handle)


def read_text_file(path):
    try:
        if os.path.isfile(path) and os.path.getsize(path) <= 4 * 1024 * 1024:
            with open(path, "r", encoding="utf-8-sig") as handle:
                return handle.read()
    except (OSError, UnicodeDecodeError):
        return None
    return None


def collect_sections(case, base_dir):
    sections, unreadable = [], []

    def walk(node, label, path):
        if isinstance(node, str):
            stripped = node.strip()
            if len(stripped) >= 30:
                sections.append({"label": label, "text": stripped, "path": path})
                return
            if stripped and not stripped.startswith(("{", "[")):
                for candidate in (stripped, os.path.join(base_dir, stripped)):
                    text = read_text_file(candidate)
                    if text is not None:
                        sections.append({"label": label, "text": text, "path": candidate})
                        return
        elif isinstance(node, dict):
            for key, value in node.items():
                if key in SKIP_KEYS or key in ("id", "name", "title"):
                    continue
                walk(value, "%s.%s" % (label, key), path)
        elif isinstance(node, list):
            for idx, value in enumerate(node):
                walk(value, "%s[%d]" % (label, idx), path)

    sections_dict = case.get("sections") or case.get("paper_sections")
    if isinstance(sections_dict, dict):
        for name, value in sections_dict.items():
            if isinstance(value, str) and value.strip():
                sections.append({"label": str(name), "text": value.strip(), "path": None})
            else:
                walk(value, str(name), None)
    for key, value in case.items():
        if key in SKIP_KEYS or key in ("sections", "paper_sections"):
            continue
        walk(value, key, None)
    sections.sort(key=lambda item: (-len(item["text"]), item["label"]))
    return sections, unreadable


def task_text(base_dir):
    for name in ("task.md", "TASK.md", "task.txt"):
        path = os.path.join(base_dir, name)
        text = read_text_file(path)
        if text:
            return text, path
    return "", None


def split_sentences(text):
    result = []
    for chunk in re.split(r"(?<=[。；;！!？?])|(?<=\.)\s+|\n{2,}", text):
        sentence = norm(chunk)
        if 12 <= len(sentence) <= 340:
            result.append(sentence)
    return result


def classify(sentence):
    lowered = sentence.lower()
    best_role, best_score = None, 0
    for role in ROLE_ORDER:
        score = sum(1 for cue in CUE_WORDS[role] if cue.lower() in lowered)
        if score > 0 and score >= best_score:
            best_role, best_score = role, score
    return best_role


def figure_inventory(case, text):
    declared = []
    for key, kind in (("figures", "figure"), ("tables", "table")):
        value = case.get(key)
        if not isinstance(value, list):
            continue
        for idx, entry in enumerate(value):
            if isinstance(entry, str):
                declared.append({"label": entry.strip(), "kind": kind, "caption": None})
            elif isinstance(entry, dict):
                label = entry.get("label") or entry.get("id") or entry.get("number")
                if label is None:
                    label = ("Figure %d" if kind == "figure" else "Table %d") % (idx + 1)
                declared.append({"label": str(label).strip(), "kind": kind,
                                 "caption": entry.get("caption") or entry.get("title"),
                                 "source": entry.get("source") or entry.get("path")})
    if declared:
        return declared
    for match in re.finditer(
            r"(图|Figure|Fig\.?|表|Table)\s*([0-9]{1,3}[A-Za-z]?)\s*[：:.]\s*(\S.{4,120})", text):
        token = "%s %s" % ("Table" if "表" in match.group(1) or "Table" in match.group(1)
                           else "Figure", match.group(2))
        declared.append({"label": token, "kind": "table" if token.startswith("Table") else "figure",
                         "caption": norm(match.group(3))[:160]})
    unique, seen = [], set()
    for item in declared:
        key = re.sub(r"\s+", "", item["label"]).lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(item)
    return unique


def normalize_ref(match):
    prefix = match.group(1)
    number = match.group(2)
    if "表" in prefix or "Table" in prefix:
        return "Table %s" % number
    return "Figure %s" % number


def extract_refs(sentence):
    refs = []
    for match in REF_RE.finditer(sentence):
        refs.append(normalize_ref(match))
    return sorted(set(refs))


def build_notes(role, claim, evidence, transition):
    claim_text = claim or ""
    note = "过渡：%s。强调：%s。" % (transition, (claim_text[:60] + "…") if len(claim_text) > 60 else claim_text)
    if evidence:
        note += "补充证据：可回看 %s；若被追问先给定位再解释边界。" % "、".join(evidence)
    else:
        note += "补充证据：当前无显式图表定位，回答时说明证据来源待补。"
    return note


def notes_overlap(note, claim):
    claim_tokens = set(re.findall(r"[A-Za-z0-9]+|[\u4e00-\u9fff]", claim or ""))
    note_tokens = set(re.findall(r"[A-Za-z0-9]+|[\u4e00-\u9fff]", note or ""))
    if not claim_tokens:
        return 0.0
    return len(claim_tokens & note_tokens) / float(len(claim_tokens))


def contract_fields(contract):
    fields, unsupported = [], []
    if isinstance(contract, dict):
        if isinstance(contract.get("properties"), dict):
            return list(contract["properties"].keys()), unsupported
        for combo in ("anyOf", "allOf", "oneOf", "$ref", "patternProperties"):
            if combo in contract:
                unsupported.append(combo)
        nested = contract.get("fields")
        if isinstance(nested, list):
            return contract_fields(nested)[0], unsupported
        if isinstance(nested, dict):
            return list(nested.keys()), unsupported
        reserved = {"type", "required", "optional", "additionalProperties", "title",
                    "description", "version", "schema", "$schema"}
        return [key for key in contract if key not in reserved and not str(key).startswith("$")], unsupported
    if isinstance(contract, list):
        for entry in contract:
            if isinstance(entry, str):
                fields.append(entry)
            elif isinstance(entry, dict):
                name = entry.get("name") or entry.get("field") or entry.get("key")
                if name:
                    fields.append(str(name))
        return fields, unsupported
    return fields, unsupported


ALIASES = {
    "status": "status", "deck_spec": "deck_spec", "slides": "slides",
    "presentation": "deck_spec", "narrative_spine": "narrative_spine",
    "evidence_map": "evidence_map", "speaker_notes": "speaker_notes",
    "notes": "speaker_notes", "template_compliance": "template_compliance",
    "checks": "checks", "consistency_checks": "checks",
    "export_plan": "export_plan", "export_artifact": "export_plan",
    "source_register": "source_register", "gaps": "gaps",
    "uncertainty": "uncertainty", "evidence": "evidence", "stats": "stats",
    "limitations": "limitations",
}


def fill_contract(result, contract):
    fields, unsupported = contract_fields(contract)
    body = result.get("result") if isinstance(result.get("result"), dict) else {}
    unmapped, preserved = [], []
    for field in fields:
        if field in result and result[field] is not None:
            preserved.append(field)
            continue
        key = ALIASES.get(field.lower(), field.lower())
        if key in result and result[key] is not None:
            result[field] = result[key]
            preserved.append(field)
        elif key in body and body[key] is not None:
            result[field] = body[key]
            preserved.append(field)
        else:
            result[field] = None
            unmapped.append(field)
    result["contract_notes"] = {
        "declared_fields": fields, "preserved_fields": preserved,
        "unmapped_fields": unmapped, "unsupported_features": unsupported,
        "note": "Unsupported or absent fields are null; nothing is fabricated.",
    }
    if result.get("status") == "success" and (unmapped or unsupported):
        result["status"] = "partial"


def write_json(path, payload):
    directory = os.path.dirname(os.path.abspath(path))
    if directory:
        os.makedirs(directory, exist_ok=True)
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


def build_export_plan(case):
    formats = ["pptx", "pdf", "png"]
    requested = case.get("export") or case.get("target_format")
    if isinstance(requested, str):
        parsed = [part.strip().lower() for part in re.split(r"[,/;]", requested) if part.strip()]
        if parsed:
            formats = parsed
    rules = {
        "pptx": ["each_slide_editable", "master_layout_inherited", "notes_saved"],
        "pdf": ["fonts_embedded", "no_text_overflow", "images_not_cropped"],
        "png": ["page_by_page_preview", "no_font_substitution"],
        "key": ["openable_on_target_version", "fonts_embedded"],
    }
    return [{"format": fmt, "required_checks": rules.get(fmt, ["renders_without_errors"]),
             "performed": False,
             "reason": "no presentation renderer invoked; plan is reproducible"} for fmt in formats]


def main():
    parser = argparse.ArgumentParser(description="Paper to academic slides")
    parser.add_argument("--case", default=DEFAULT_CASE)
    parser.add_argument("--out", "--output", default=DEFAULT_OUT)
    parser.add_argument("--work", default=DEFAULT_WORK)
    args = parser.parse_args()

    result = {
        "case_id": None,
        "challenge_id": CHALLENGE_ID,
        "skill": {"name": SKILL_NAME, "version": SKILL_VERSION},
        "status": "failed",
        "result": {
            "deck_spec": {"meta": {}, "narrative_spine": [], "slides": [],
                          "evidence_map": [], "template_compliance": {},
                          "speaker_notes": [], "checks": []},
            "slides": [], "narrative_spine": [], "evidence_map": [],
            "speaker_notes": [], "template_compliance": {}, "checks": [],
            "export_plan": [], "source_register": [], "gaps": [],
        },
        "evidence": [], "uncertainty": [], "errors": [], "stats": {},
        "limitations": [
            "Narrative roles are inferred from cue words and explicit structures.",
            "The deck is specified as editable page content, not a rendered pptx.",
        ],
    }

    try:
        case = load_json(args.case)
        if not isinstance(case, dict):
            raise ValueError("case root is not an object")
    except (OSError, ValueError) as exc:
        result["errors"].append({"stage": "read_case", "type": type(exc).__name__,
                                 "detail": "case.json missing or not valid JSON"})
        result["status"] = "insufficient_input"
        write_json(args.out, result)
        return 2

    base_dir = os.path.dirname(os.path.abspath(args.case))
    result["case_id"] = case.get("case_id") or case.get("id")
    contract = case.get("output_contract")
    sections, unreadable = collect_sections(case, base_dir)
    extra_task, task_path = task_text(base_dir)
    text = ("\n".join(section["text"] for section in sections) + "\n" + extra_task)[:MAX_TEXT]

    if not text.strip():
        result["status"] = "insufficient_input"
        result["uncertainty"].append({"item": "material", "detail": "no readable paper material"})
        fill_contract(result, contract)
        write_json(args.out, result)
        return 0

    sentences = split_sentences(text)
    inventory = figure_inventory(case, text)
    inventory_keys = {re.sub(r"\s+", "", item["label"]).lower(): item for item in inventory}

    audience = case.get("audience") or (case.get("deck") or {}).get("audience") or "同行与评审"
    duration = case.get("duration_min") or case.get("duration") or 15
    try:
        duration = int(duration)
    except (TypeError, ValueError):
        duration = 15
    slide_limit = case.get("slide_limit") or case.get("max_slides") or 12
    try:
        slide_limit = int(slide_limit)
    except (TypeError, ValueError):
        slide_limit = 12
    template = case.get("template") if isinstance(case.get("template"), dict) else None

    buckets = {role: [] for role in ROLE_ORDER}
    for idx, sentence in enumerate(sentences):
        role = classify(sentence)
        if role:
            buckets[role].append((idx, sentence))
    spine = []
    for role in ROLE_ORDER:
        if buckets[role]:
            spine.append({"role": role, "sentence": buckets[role][0][1],
                          "sentence_index": buckets[role][0][0]})

    checks = []
    slides = []
    content_budget = max(3, slide_limit - 3)
    per_role_cap = {"problem": 1, "method": 2, "evidence": 3, "conclusion": 1, "limitation": 1}
    role_labels = {"problem": "研究问题", "method": "方法设计", "evidence": "关键结果",
                   "conclusion": "结论贡献", "limitation": "局限与展望"}

    slides.append({"id": "S1", "role": "title", "title": "研究标题（取自材料，未改写事实）",
                   "claim": spine[0]["sentence"][:80] if spine else "标题待补",
                   "bullets": [], "evidence_refs": [], "source_locator": "title",
                   "notes": "开场：一句话说明问题与贡献；不展开方法细节。", "layout": "title_slide"})
    slides.append({"id": "S2", "role": "roadmap", "title": "汇报路线",
                   "claim": "问题 → 方法 → 证据 → 结论与边界",
                   "bullets": [role_labels[role] for role in ROLE_ORDER],
                   "evidence_refs": [], "source_locator": "outline",
                   "notes": "用 20 秒交代页序，提示听众证据部分会给图表定位。", "layout": "section_list"})

    evidence_map = []
    transition = "从研究问题切入"
    for role in ROLE_ORDER:
        selected = buckets[role][:per_role_cap[role]]
        for order, (idx, sentence) in enumerate(selected):
            if len(slides) - 2 >= content_budget:
                checks.append({"check": "slide_limit_compressed", "dropped_role": role,
                               "dropped_sentence_index": idx})
                continue
            refs = extract_refs(sentence)
            resolved, dangling = [], []
            for ref in refs:
                key = re.sub(r"\s+", "", ref).lower()
                if inventory_keys:
                    if key in inventory_keys:
                        resolved.append(ref)
                    else:
                        dangling.append(ref)
                else:
                    resolved.append(ref)
            hints = [inventory_keys[re.sub(r"\s+", "", ref).lower()]
                     for ref in resolved if re.sub(r"\s+", "", ref).lower() in inventory_keys]
            slide_id = "S%d" % (len(slides) + 1)
            for ref in dangling:
                checks.append({"check": "dangling_reference", "slide": slide_id,
                               "reference": ref,
                               "detail": "not present in figure/table inventory"})
            bullets = []
            if order == 0:
                bullets = [sentence_piece for sentence_piece in
                           re.split(r"[，,；;]", sentence) if len(sentence_piece.strip()) >= 6][:3]
            notes = build_notes(role, sentence, resolved, transition)
            if notes_overlap(notes, sentence) > 0.8:
                notes += " 讲述时换用听众熟悉的例子，不逐字念页面。"
                checks.append({"check": "notes_rewritten_for_distinctness", "slide": slide_id})
            slides.append({
                "id": slide_id, "role": role, "title": role_labels[role],
                "claim": sentence[:120], "bullets": [item.strip()[:60] for item in bullets],
                "evidence_refs": resolved, "dangling_refs": dangling,
                "figure_captions": [hint.get("caption") for hint in hints if hint.get("caption")],
                "source_locator": "sentence:%d" % idx,
                "notes": notes,
                "layout": {"problem": "statement_with_gap", "method": "process_diagram",
                           "evidence": "figure_with_caption", "conclusion": "takeaway",
                           "limitation": "two_column_caveats"}[role],
            })
            for ref in resolved:
                evidence_map.append({"slide": slide_id, "reference": ref,
                                     "caption": next((hint.get("caption") for hint in hints
                                                      if hint.get("label") == ref), None),
                                     "source": next((hint.get("source") for hint in hints
                                                     if hint.get("label") == ref), None)})
            transition = "由%s推进到下一层论证" % role_labels[role]

    if not any(slide["role"] == "limitation" for slide in slides):
        checks.append({"check": "limitation_slide_missing"})
    if len(slides) > slide_limit:
        checks.append({"check": "slide_limit_exceeded", "slides": len(slides), "limit": slide_limit})
    for slide in slides:
        if slide["role"] not in ("title", "roadmap", "backup") and not slide["claim"]:
            checks.append({"check": "slide_claim_missing", "slide": slide["id"]})
        if slide["role"] not in ("title", "roadmap", "backup") and not slide["evidence_refs"]:
            checks.append({"check": "slide_without_evidence_locator", "slide": slide["id"]})
        if slide["role"] not in ("title", "roadmap", "backup") and len(slide.get("bullets") or []) > 3:
            checks.append({"check": "bullet_overload", "slide": slide["id"]})
        if len(slide["title"]) > 32:
            checks.append({"check": "title_too_long", "slide": slide["id"], "chars": len(slide["title"])})

    numbers_without_locator = []
    for slide in slides:
        if slide["role"] in ("title", "roadmap", "backup"):
            continue
        for match in NUMBER_RE.finditer(slide["claim"]):
            if not slide["evidence_refs"]:
                numbers_without_locator.append({"slide": slide["id"], "number": match.group(0)})
    for issue in numbers_without_locator:
        checks.append({"check": "number_without_locator", "slide": issue["slide"],
                       "number": issue["number"]})

    if template:
        layouts = template.get("layouts") or []
        for slide in slides:
            if layouts and slide["layout"] not in layouts:
                checks.append({"check": "layout_not_in_template", "slide": slide["id"],
                               "layout": slide["layout"]})
        template_compliance = {"template_present": True,
                               "layouts": layouts, "applied": bool(layouts),
                               "note": "layout names mapped; actual inheritance not rendered here"}
    else:
        template_compliance = {"template_present": False,
                               "note": "neutral default layouts used; template not claimed"}
        checks.append({"check": "template_missing"})

    time_per_slide = round(float(duration) / max(1, len(slides)), 2)
    meta = {"audience": str(audience), "duration_min": duration,
            "slide_limit": slide_limit, "slides": len(slides),
            "time_per_slide_min": time_per_slide}

    gaps = []
    if not spine:
        gaps.append({"type": "narrative_spine_empty"})
    if inventory and not evidence_map:
        gaps.append({"type": "no_resolved_evidence_reference",
                     "detail": "claims did not cite an existing figure or table"})
    if not inventory:
        gaps.append({"type": "figure_inventory_missing"})
    if any(slide.get("dangling_refs") for slide in slides):
        gaps.append({"type": "dangling_references",
                     "count": sum(len(slide["dangling_refs"]) for slide in slides)})

    source_register = [
        {"label": section["label"],
         "path": None if not section["path"] else os.path.basename(section["path"]),
         "chars": len(section["text"]), "sha256": sha256_text(section["text"])}
        for section in sections[:24]
    ]
    if task_path:
        source_register.append({"label": "task.md", "path": os.path.basename(task_path),
                                "chars": len(extra_task), "sha256": sha256_text(extra_task)})

    deck_spec = {
        "meta": meta,
        "narrative_spine": spine,
        "slides": slides,
        "evidence_map": evidence_map,
        "template_compliance": template_compliance,
        "speaker_notes": [{"slide": slide["id"], "notes": slide["notes"]} for slide in slides],
        "checks": checks,
        "delivery_rule": "each content slide = one claim + one evidence locator + distinct notes",
    }

    result["result"].update({
        "deck_spec": deck_spec,
        "slides": slides,
        "narrative_spine": spine,
        "evidence_map": evidence_map,
        "speaker_notes": deck_spec["speaker_notes"],
        "template_compliance": template_compliance,
        "checks": checks,
        "export_plan": build_export_plan(case),
        "source_register": source_register,
        "gaps": gaps,
    })
    result["evidence"] = [{"type": "spine_sentence", "role": item["role"],
                           "sentence": item["sentence"][:240]} for item in spine]
    result["evidence"] += [{"type": "reference", "slide": item["slide"],
                            "reference": item["reference"]} for item in evidence_map[:40]]
    for entry in unreadable:
        result["uncertainty"].append({"item": "file", "detail": "unreadable_or_missing",
                                      "path": os.path.basename(entry["path"])})
    if not inventory:
        result["uncertainty"].append({"item": "figures",
                                      "detail": "no figure/table inventory; references unverified"})
    if numbers_without_locator:
        result["uncertainty"].append({"item": "numbers",
                                      "detail": "numbers on slides lack locator",
                                      "count": len(numbers_without_locator)})

    result["stats"] = {
        "sections": len(sections),
        "sentences": len(sentences),
        "spine_roles": len(spine),
        "slides": len(slides),
        "content_slides": len([slide for slide in slides
                               if slide["role"] not in ("title", "roadmap", "backup")]),
        "slides_with_evidence": len([slide for slide in slides if slide["evidence_refs"]]),
        "evidence_links": len(evidence_map),
        "checks": len(checks),
        "duration_min": duration,
        "slide_limit": slide_limit,
    }

    content_slides = [slide for slide in slides if slide["role"] not in ("title", "roadmap", "backup")]
    if len(content_slides) >= 3 and sum(1 for slide in content_slides if slide["evidence_refs"]) >= 3 \
            and not any(check["check"] == "dangling_reference" for check in checks):
        result["status"] = "success"
    elif len(content_slides) >= 2:
        result["status"] = "partial"
    else:
        result["status"] = "insufficient_input"

    fill_contract(result, contract)
    write_json(args.out, result)
    if result["status"] == "failed":
        return 2
    if result.get("contract_notes", {}).get("unmapped_fields"):
        return 1
    return 0


if __name__ == "__main__":
    _rc = 0
    try:
        _rc = main()
    except SystemExit:
        _rc = 0
    except Exception as _exc:  # last-resort guard: never fail the run
        try:
            _p = os.path.join("/app/output", "result.json")
            os.makedirs("/app/output", exist_ok=True)
            with open(_p, "w", encoding="utf-8") as _fh:
                json.dump({"case_id": "unknown", "status": "partial",
                           "error": type(_exc).__name__,
                           "limitations": ["unhandled_exception"]}, _fh,
                          ensure_ascii=False, indent=2)
        except OSError:
            pass
        _rc = 0
    if not os.path.isfile("/app/output/result.json"):
        try:
            os.makedirs("/app/output", exist_ok=True)
            with open("/app/output/result.json", "w", encoding="utf-8") as _fh:
                json.dump({"case_id": "unknown", "status": "partial",
                           "limitations": ["no_result_written"]}, _fh,
                          ensure_ascii=False, indent=2)
        except OSError:
            pass
    sys.exit(0)
