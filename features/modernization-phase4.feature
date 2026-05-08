Feature: Modernization Phase 4 — shape-tree ergonomics
  In order to traverse, look up, and inspect shapes ergonomically
  As a developer using python-pptx
  I need iter_leaf_shapes, mapping-like name access, find_by_xpath, and selection-pane ordering


  Scenario: Mapping-like name access on a slide's shapes
    Given a fresh slide with a title placeholder
     Then shapes["Title 1"] returns the title shape
      And "Title 1" is in shapes
      And "Bogus" is not in shapes


  Scenario: Mapping-like name access on a slide's placeholders
    Given a fresh slide with a title placeholder
     Then placeholders["Title 1"] returns the title placeholder
      And "Title 1" is in placeholders


  Scenario: shapes.keys() returns the list of shape names
    Given a fresh slide with a title placeholder
     Then shapes.keys() includes "Title 1"


  Scenario: in_selection_pane_order reverses XML order
    Given a fresh slide with a title placeholder
     Then shapes.in_selection_pane_order() reverses iteration order


  Scenario: iter_leaf_shapes yields top-level shapes when no groups present
    Given a fresh slide with a title placeholder
     Then iter_leaf_shapes() yields the same shapes as iteration


  Scenario: find_by_xpath returns a non-empty list for a known element
    Given a fresh slide with a title placeholder
     Then title.find_by_xpath(".//p:nvSpPr") has length 1


  Scenario: find_by_xpath returns empty list on no match
    Given a fresh slide with a title placeholder
     Then title.find_by_xpath(".//a:nope_no_match") is empty
