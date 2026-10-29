from pathlib import Path
import re

import pandas as pd
import fitz
from docx import Document


QUESTION_PREFIX_RE = re.compile(
    r"^\s*(?:Q(?:UESTION)?\s*\d+\s*[.)\-:]?\s*)(.*)$",
    re.IGNORECASE,
)
NUMBERED_RE = re.compile(r"^\s*(\d+)\s*[.)\-:]\s+(.+?)\s*$")
OPTION_RE = re.compile(r"^\s*(?:[-•*]|[A-Za-z][.)]|\d+[.)])\s*(.+?)\s*$")


def extract_text(path: Path) -> str:
    """Extract selectable text from PDF, DOCX or TXT. OCR is not included."""
    path = Path(path)
    ext = path.suffix.lower()
    if ext == ".pdf":
        with fitz.open(path) as doc:
            text = "\n".join(page.get_text("text") for page in doc)
        if not text.strip():
            raise ValueError(
                "Aucun texte extractible dans ce PDF. Il semble numérisé : "
                "effectuez un OCR local puis réessayez."
            )
        return text
    if ext == ".docx":
        d = Document(path)
        parts = [p.text for p in d.paragraphs if p.text.strip()]
        for table in d.tables:
            for row in table.rows:
                parts.append(" | ".join(cell.text for cell in row.cells))
        return "\n".join(parts)
    if ext == ".txt":
        return path.read_text(encoding="utf-8", errors="ignore")
    raise ValueError("Formats acceptés : PDF, DOCX ou TXT.")


def _clean_line(line: str) -> str:
    return re.sub(r"\s+", " ", line).strip()


def parse_questionnaire(text: str):
    """Conservatively parse questions and answer options from extracted text."""
    lines = [_clean_line(line) for line in text.splitlines() if line.strip()]
    items = []
    current = None

    def flush():
        nonlocal current
        if current is not None:
            current["choices"] = list(dict.fromkeys(c for c in current["choices"] if c))
            items.append(current)
            current = None

    for line in lines:
        qprefix = QUESTION_PREFIX_RE.match(line)
        numbered = NUMBERED_RE.match(line)
        is_question = False
        label = line

        if qprefix:
            is_question = True
            label = qprefix.group(1).strip()
        elif numbered and ("?" in line or current is None):
            is_question = True
            label = numbered.group(2).strip()
        elif "?" in line and not (current is not None and OPTION_RE.match(line)):
            is_question = True
            label = re.sub(r"^\s*\d+\s*[.)\-:]\s*", "", line).strip()

        if is_question and label:
            flush()
            current = {"label": label, "choices": [], "hint": ""}
            continue

        option = OPTION_RE.match(line)
        if option and current is not None:
            current["choices"].append(option.group(1).strip())
            continue

        if current is not None:
            current["hint"] = (current["hint"] + " " + line).strip()

    flush()
    return items


def make_name(label, idx):
    base = re.sub(r"[^a-z0-9_]+", "_", label.lower()).strip("_")
    base = re.sub(r"_+", "_", base)[:40]
    if not base or not base[0].isalpha():
        base = f"q{idx}_{base}"
    return base


def to_xlsform(items, out: Path, title="Questionnaire converti", form_id="questionnaire_converti"):
    survey = []
    choices = []
    seen = set()

    for i, item in enumerate(items, 1):
        label = str(item.get("label", "")).strip()
        if not label:
            continue

        name = make_name(label, i)
        original = name
        suffix = 2
        while name in seen:
            name = f"{original}_{suffix}"
            suffix += 1
        seen.add(name)

        options = [str(option).strip() for option in item.get("choices", []) if str(option).strip()]
        if options:
            list_name = f"list_{name}"[:45]
            qtype = f"select_one {list_name}"
            option_names = set()
            for k, option_label in enumerate(options, 1):
                option_name = make_name(option_label, k)[:40] or f"choice_{k}"
                base_option_name = option_name
                option_suffix = 2
                while option_name in option_names:
                    option_name = f"{base_option_name}_{option_suffix}"
                    option_suffix += 1
                option_names.add(option_name)
                choices.append({"list_name": list_name, "name": option_name, "label": option_label})
        else:
            qtype = "text"

        survey.append({
            "type": qtype, "name": name, "label": label,
            "hint": item.get("hint", ""), "required": "", "relevant": "",
            "constraint": "", "constraint_message": "", "calculation": "",
            "appearance": "", "choice_filter": "",
        })

    if not survey:
        raise ValueError("Aucune question exploitable détectée ; XLSForm non généré.")

    settings = [{
        "form_title": title,
        "form_id": re.sub(r"[^A-Za-z0-9_-]", "_", form_id)[:60] or "questionnaire_converti",
        "version": "1",
        "default_language": "French (fr)",
    }]
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame(survey).to_excel(writer, sheet_name="survey", index=False)
        pd.DataFrame(choices, columns=["list_name", "name", "label"]).to_excel(writer, sheet_name="choices", index=False)
        pd.DataFrame(settings).to_excel(writer, sheet_name="settings", index=False)
    return {"questions": len(survey), "choices": len(choices), "output": str(out)}


def validate_items(items):
    warnings = []
    if not items:
        warnings.append("Aucune question détectée automatiquement. Revue humaine obligatoire.")
    seen_names = set()
    for i, item in enumerate(items, 1):
        label = str(item.get("label", "")).strip()
        if len(label) < 3:
            warnings.append(f"Question {i}: libellé trop court.")
        key = make_name(label, i)
        if key in seen_names:
            warnings.append(f"Question {i}: nom de variable potentiellement dupliqué.")
        seen_names.add(key)
        if item.get("choices") and len(item["choices"]) < 2:
            warnings.append(f"Question {i}: une seule modalité détectée ; vérifiez les choix.")
    return {"valid": len(items) > 0, "warnings": warnings}
