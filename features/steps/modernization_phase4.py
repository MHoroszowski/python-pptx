"""Gherkin steps for Modernization Phase 4 (issue #29) — shape-tree ergonomics."""

from __future__ import annotations

from behave import then


# the "Given a fresh slide with a title placeholder" step is shared from
# features/steps/modernization_phase2.py (Phase 2)


# then ====================================================


@then('shapes["Title 1"] returns the title shape')
def then_shapes_str_key_returns_title(context):
    sh = context.slide.shapes["Title 1"]
    assert sh.name == "Title 1", sh.name


@then('"Title 1" is in shapes')
def then_title_in_shapes(context):
    assert "Title 1" in context.slide.shapes


@then('"Bogus" is not in shapes')
def then_bogus_not_in_shapes(context):
    assert "Bogus" not in context.slide.shapes


@then('placeholders["Title 1"] returns the title placeholder')
def then_placeholders_str_key(context):
    ph = context.slide.placeholders["Title 1"]
    assert ph.name == "Title 1", ph.name


@then('"Title 1" is in placeholders')
def then_title_in_placeholders(context):
    assert "Title 1" in context.slide.placeholders


@then('shapes.keys() includes "Title 1"')
def then_shapes_keys_includes_title(context):
    assert "Title 1" in context.slide.shapes.keys()


@then("shapes.in_selection_pane_order() reverses iteration order")
def then_selection_pane_reverses(context):
    xml_order = [s.name for s in context.slide.shapes]
    sp_order = [s.name for s in context.slide.shapes.in_selection_pane_order()]
    assert sp_order == list(reversed(xml_order)), (sp_order, xml_order)


@then("iter_leaf_shapes() yields the same shapes as iteration")
def then_iter_leaf_matches_iter(context):
    leaves = [s.name for s in context.slide.shapes.iter_leaf_shapes()]
    top = [s.name for s in context.slide.shapes]
    assert leaves == top, (leaves, top)


@then('title.find_by_xpath(".//p:nvSpPr") has length 1')
def then_xpath_match_length_1(context):
    title = context.slide.shapes.title
    results = title.find_by_xpath(".//p:nvSpPr")
    assert len(results) == 1, len(results)


@then('title.find_by_xpath(".//a:nope_no_match") is empty')
def then_xpath_empty(context):
    title = context.slide.shapes.title
    assert title.find_by_xpath(".//a:nope_no_match") == []
