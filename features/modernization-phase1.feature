Feature: Modernization Phase 1 — PathLike, PERCENT_40, slide.background.element
  In order to use python-pptx with modern Python idioms
  As a developer using python-pptx
  I need pathlib.Path support, the PERCENT_40 enum typo fixed,
  and slide.background.element to return the actual <p:bg> element


  Scenario: Open a presentation from a pathlib.Path
    Given a freshly-saved presentation at a Path
     When I call Presentation(path) with the Path
     Then I get a presentation back


  Scenario: Save a presentation to a pathlib.Path
    Given a fresh presentation
     When I save it to a Path
     Then a non-empty .pptx file exists at that Path


  Scenario: PERCENT_40 enum is exposed with the correct name
    Then MSO_PATTERN_TYPE.PERCENT_40 exists with xml_value pct40
     And the broken name ERCENT_40 does not exist


  Scenario: slide.background.element returns the <p:bg> element
    Given a fresh slide on a fresh presentation
     Then slide.background.element local-name is bg
