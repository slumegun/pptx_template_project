import pytest
from docx import Document
from engine.content import extract_facts, _split_statements
from engine.models import Fact
from engine.swarm import validate_slide
from engine.provider import ModelProviderError


def test_leading_numbers_and_negative_values_survive():
    for text in ["2026 год: запуск нового продукта.", "35% пользователей выбрали второй вариант.", "-12 рублей составляет изменение стоимости."]:
        assert _split_statements(text) == [text]
    assert _split_statements("1. Первый пункт с содержательным фактом.") == ["Первый пункт с содержательным фактом."]


def test_docx_tables_are_sources(tmp_path):
    path = tmp_path / "source.docx"
    doc = Document()
    row = doc.add_table(rows=1, cols=2).rows[0]
    row.cells[0].text = "Выручка продукта"
    row.cells[1].text = "125 миллионов рублей"
    doc.save(path)
    facts = extract_facts("", [path])
    assert len(facts) == 1
    assert "125" in facts[0].excerpt
    assert facts[0].location == "table 1, row 1"


def test_worker_cannot_add_numbers_or_drop_references():
    fact = Fact("f1", "Выручка выросла на 12,5%.", "brief", "line 1", "Выручка выросла на 12,5%.")
    slide = {"title": "Рост выручки", "bullets": ["Выручка выросла на 12.5%."], "fact_ids": ["f1"]}
    assert validate_slide(slide, "slide_1", {"f1": fact}).fact_ids == ["f1"]
    slide["bullets"] = ["Выручка выросла на 25%."]
    with pytest.raises(ModelProviderError, match="числа"):
        validate_slide(slide, "slide_1", {"f1": fact})
    slide["bullets"] = ["Выручка выросла."]
    with pytest.raises(ModelProviderError, match="набор фактов"):
        validate_slide(slide, "slide_1", {"f1": fact}, required_refs={"f1", "f2"})


def test_short_brief_cannot_gain_an_unmentioned_realtime_capability():
    fact = Fact("f1", "Платформа помогает отслеживать пожары на карте.", "brief", "line 1",
                "Платформа помогает отслеживать пожары на карте.")
    slide = {"title": "Мониторинг пожаров", "bullets": ["Система отслеживает пожары в реальном времени."],
             "fact_ids": ["f1"]}
    with pytest.raises(ModelProviderError, match="неподтверждённую возможность"):
        validate_slide(slide, "slide_1", {"f1": fact})
    slide["bullets"] = ["Система помогает наблюдать пожары на карте."]
    assert validate_slide(slide, "slide_1", {"f1": fact}).bullets == slide["bullets"]

def test_brief_instructions_do_not_become_product_facts():
    facts = extract_facts('Покажи редактируемую схему связей.\nРадарная ветвь использует Sentinel-1.\nСохрани все веса без искажения.')
    assert [f.excerpt for f in facts] == ['Радарная ветвь использует Sentinel-1.']
