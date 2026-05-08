Feature: Modernization & Ergonomics Phase 2 — bug fixes + by_name
  In order to inspect fonts without polluting the document, round-trip datetimes faithfully, and look up shapes by name
  As a developer using python-pptx
  I need a non-mutating Font.color getter, tz-aware core-property datetimes, and Shapes.by_name(name)


  Scenario: Reading font.color does not insert a:solidFill
    Given a fresh slide with a title placeholder
     When I read run.font.color.rgb on an unstyled run
     Then the underlying rPr XML is unchanged from before the access


  Scenario: Setting font.color.rgb materializes a:solidFill lazily
    Given a fresh slide with a title placeholder
     When I set run.font.color.rgb to RGBColor(0xFF, 0x00, 0x00)
     Then the underlying rPr now contains an a:solidFill child
      And run.font.color.rgb reads back as FF0000


  Scenario: Tz-aware core_properties.created round-trips faithfully
    Given a fresh presentation for core-property datetimes
     When I set core_properties.created to a tz-aware UTC datetime
     Then the reloaded core_properties.created is tz-aware


  Scenario: shapes.by_name returns the matching shape
    Given a fresh slide with a title placeholder
     Then shapes.by_name("Title 1") returns the title shape


  Scenario: shapes.by_name raises KeyError on miss
    Given a fresh slide with a title placeholder
     Then shapes.by_name("Bogus") raises KeyError
