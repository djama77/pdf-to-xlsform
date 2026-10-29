"""Command-line interface for the PDF-to-XLSForm engine."""
import argparse
import json
import sys
from pathlib import Path

from .engine import extract_text, parse_questionnaire, to_xlsform, validate_items


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="pdf-to-xlsform",
        description="Convert a PDF, DOCX or TXT questionnaire to an XLSForm workbook locally.",
    )
    parser.add_argument("input", type=Path, help="Input questionnaire (.pdf, .docx or .txt)")
    parser.add_argument("-o", "--output", type=Path, default=Path("questionnaire.xlsx"),
                        help="Output XLSX file (default: questionnaire.xlsx)")
    parser.add_argument("--title", default="Questionnaire converti", help="Form title")
    parser.add_argument("--form-id", default="questionnaire_converti", help="Stable XLSForm form ID")
    args = parser.parse_args(argv)

    try:
        source_text = extract_text(args.input)
        items = parse_questionnaire(source_text)
        report = validate_items(items)
        if not report["valid"]:
            print("Erreur : aucune question détectée. Aucun fichier n'a été généré.", file=sys.stderr)
            for warning in report["warnings"]:
                print(f"- {warning}", file=sys.stderr)
            return 2

        result = to_xlsform(items, args.output, title=args.title, form_id=args.form_id)
        print(json.dumps({"result": result, "validation": report}, ensure_ascii=False, indent=2))
        if report["warnings"]:
            print("\nAttention : vérifiez manuellement les questions et les choix avant déploiement.", file=sys.stderr)
        return 0
    except (OSError, ValueError) as exc:
        print(f"Erreur : {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
