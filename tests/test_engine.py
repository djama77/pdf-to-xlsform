from pathlib import Path

import fitz
import pandas as pd
import pytest

from pdf_to_kobocollect import extract_text, parse_questionnaire, to_xlsform, validate_items


def test_numbered_answer_options_stay_with_their_question(tmp_path: Path):
    text = (
        "Q1. Quel est votre sexe ?\n1. Homme\n2. Femme\n"
        "Q2. Quel est votre âge ?\n1. Moins de 18 ans\n2. 18 ans ou plus\n"
    )
    items = parse_questionnaire(text)
    assert len(items) == 2
    assert items[0]["choices"] == ["Homme", "Femme"]
    assert items[1]["choices"] == ["Moins de 18 ans", "18 ans ou plus"]
    output = tmp_path / "form.xlsx"
    result = to_xlsform(items, output)
    assert result["questions"] == 2
    assert result["choices"] == 4
    survey = pd.read_excel(output, sheet_name="survey")
    choices = pd.read_excel(output, sheet_name="choices")
    assert len(survey) == 2
    assert len(choices) == 4
    assert survey["type"].str.startswith("select_one ").all()


def test_unmarked_question_with_question_mark_starts_new_item():
    items = parse_questionnaire(
        "Quel est votre pays ?\n- Djibouti\n- France\n"
        "Quelle est votre profession ?\n- Étudiant\n- Autre"
    )
    assert len(items) == 2
    assert items[0]["choices"] == ["Djibouti", "France"]
    assert items[1]["choices"] == ["Étudiant", "Autre"]


def test_duplicate_labels_get_unique_variable_and_choice_names(tmp_path: Path):
    items = parse_questionnaire("Q1. Choisissez une option ?\n1. Autre\n2. Autre")
    output = tmp_path / "duplicates.xlsx"
    to_xlsform(items, output)
    survey = pd.read_excel(output, sheet_name="survey")
    choices = pd.read_excel(output, sheet_name="choices")
    assert survey["name"].is_unique
    assert choices["name"].is_unique


def test_empty_questionnaire_is_not_exported(tmp_path: Path):
    with pytest.raises(ValueError, match="Aucune question"):
        to_xlsform([], tmp_path / "empty.xlsx")


def test_scanned_or_blank_pdf_returns_actionable_ocr_message(tmp_path: Path):
    pdf_path = tmp_path / "blank.pdf"
    with fitz.open() as document:
        document.new_page()
        document.save(pdf_path)
    with pytest.raises(ValueError, match="OCR local"):
        extract_text(pdf_path)


def test_validation_warns_when_only_one_choice_is_detected():
    result = validate_items([{"label": "Quel est votre statut ?", "choices": ["Étudiant"]}])
    assert result["valid"] is True
    assert any("une seule modalité" in warning for warning in result["warnings"])
