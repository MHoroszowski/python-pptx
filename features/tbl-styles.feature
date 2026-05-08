Feature: Table style API — apply built-in PowerPoint table styles
  In order to render tables in the chosen built-in PowerPoint style
  As a developer using python-pptx
  I need to read, set, and clear a table's style id by name or GUID


  Scenario: A newly added table reports the default style
    Given a 2x2 table on a fresh slide
     Then table.style_id is "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
      And table.style_name is "Medium Style 2 - Accent 1"


  Scenario: Apply a style by friendly name
    Given a 2x2 table on a fresh slide
     When I call table.apply_style("Medium Style 2 - Accent 3")
     Then table.style_id is "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}"
      And table.style_name is "Medium Style 2 - Accent 3"


  Scenario: Apply a style by raw GUID
    Given a 2x2 table on a fresh slide
     When I call table.apply_style("{2D5ABB26-0587-4C30-8999-92F81FD0307C}")
     Then table.style_id is "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
      And table.style_name is "No Style, No Grid"


  Scenario: Apply by name is case-insensitive
    Given a 2x2 table on a fresh slide
     When I call table.apply_style("light style 2 - accent 4")
     Then table.style_id is "{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}"


  Scenario: Apply an unknown name raises ValueError
    Given a 2x2 table on a fresh slide
     Then calling table.apply_style("Bogus Name") raises ValueError


  Scenario: Clear the style by setting style_id to None
    Given a 2x2 table on a fresh slide
     When I set table.style_id to None
     Then table.style_id is None
      And table.style_name is None


  Scenario: Round-trip preserves style_id through save/reload
    Given a 2x2 table on a fresh slide
     When I call table.apply_style("Light Style 2 - Accent 4")
      And I save and reload the presentation via stream
     Then the reloaded table has style_id "{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}"
      And the reloaded table has style_name "Light Style 2 - Accent 4"
