import io

import pandas as pd
from openpyxl import load_workbook

import app


ERGON_HEADER = ["NUMFUNC", "NUMVINC", "NOME", "LOTACAO", "CARGO", "SITUACAO"]
INTRANET_HEADER = ["ESCALA", "NUMFUNC", "NUMVINC", "SERVIDOR", "SETOR", "CARGO"]


def test_make_key_ignores_formatting_and_leading_zeroes():
    assert app.make_key("001.234.567", "01.0") == "1234567-1"


def test_parse_and_clean_create_same_compound_key():
    ergon_raw = pd.DataFrame([
        ERGON_HEADER,
        ["001234567", "01", "Ana", "Hospital de Referencia de Araguaina", "Enfermeira", "ATIVO"],
    ])
    intranet_raw = pd.DataFrame([
        INTRANET_HEADER,
        ["10", "1234567", "1", "ANA", "HOSPITAL DE REFERENCIA DE ARAGUAINA", "ENFERMEIRA"],
    ])

    ergon = app.clean_ergon(app.parse_raw(ergon_raw, app.ERGON_ALIASES))
    intranet = app.clean_intranet(app.parse_raw(intranet_raw, app.INTRANET_ALIASES))

    assert ergon.loc[0, "CHAVE_VINCULO"] == "1234567-1"
    assert intranet.loc[0, "CHAVE_VINCULO"] == "1234567-1"


def test_compare_uses_hospital_and_accepts_any_duplicate_scale():
    ergon_raw = pd.DataFrame([
        ERGON_HEADER,
        ["100", "1", "Escalado", "Hospital de Referencia de Araguaina", "Técnico", "ATIVO"],
        ["200", "2", "Não escalado", "Hospital de Referencia de Araguaina", "Enfermeiro", "ATIVO"],
    ])
    intranet_raw = pd.DataFrame([
        INTRANET_HEADER,
        ["10", "100", "1", "Escalado", "HOSPITAL DE REFERENCIA DE ARAGUAINA", "Técnico"],
        ["11", "100", "1", "Escalado", "HOSPITAL DE REFERENCIA DE ARAGUAINA", "Técnico"],
    ])

    result = app.compare_datasets(
        app.clean_ergon(app.parse_raw(ergon_raw, app.ERGON_ALIASES)),
        app.clean_intranet(app.parse_raw(intranet_raw, app.INTRANET_ALIASES)),
    )

    assert len(result["scaled"]) == 1
    assert result["scaled"].iloc[0]["NOME"] == "Escalado"
    assert result["not_scaled"]["NOME"].tolist() == ["Não escalado"]


def test_unmapped_hospital_is_not_counted_as_not_scaled():
    ergon_raw = pd.DataFrame([
        ERGON_HEADER,
        ["300", "1", "Fora", "Hospital Desconhecido", "Técnico", "ATIVO"],
    ])
    intranet_raw = pd.DataFrame([INTRANET_HEADER])

    result = app.compare_datasets(
        app.clean_ergon(app.parse_raw(ergon_raw, app.ERGON_ALIASES)),
        app.clean_intranet(app.parse_raw(intranet_raw, app.INTRANET_ALIASES)),
    )

    assert result["not_scaled"].empty
    assert len(result["unmapped"]) == 1


def test_filter_and_excel_export():
    source = pd.DataFrame({
        "HOSPITAL_ERGON": ["Hospital A", "Hospital B"],
        "HOSPITAL_INTRANET": ["A", "B"],
        "CHAVE_VINCULO": ["1-1", "2-1"],
        "ENCONTRADO_INTRANET": [False, False],
    })
    filtered = app.filter_results(source, "Hospital B")
    content = app.df_to_excel_bytes({"Nao escalados": filtered})
    sheet = load_workbook(io.BytesIO(content))["Nao escalados"]

    assert filtered["CHAVE_VINCULO"].tolist() == ["2-1"]
    assert sheet.page_setup.orientation == "landscape"
    assert str(sheet.page_setup.paperSize) == str(sheet.PAPERSIZE_A4)
    assert sheet.page_setup.fitToWidth == 1
    assert sheet.page_setup.fitToHeight == 0
