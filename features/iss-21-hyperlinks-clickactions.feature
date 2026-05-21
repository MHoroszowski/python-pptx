Feature: Hyperlinks 2.0 & Click Actions (issue #21)
  In order to fully author hyperlink and click/hover behaviors
  As a developer using python-pptx
  I need ScreenTips, colors, run/macro/program/sound actions, hover, and jumps

  Scenario: Set a ScreenTip on a run hyperlink
    Given a run with an external hyperlink
     When I set the run hyperlink tooltip to "Click for details"
     Then the reopened run hyperlink tooltip is "Click for details"

  Scenario: Shape alt-text round-trips through save and reopen
    Given a shape
     When I set the shape alt-text to "Accessible description"
     Then the reopened shape alt-text is "Accessible description"

  Scenario: A tooltip-only run hyperlink round-trips at the XML layer
    Given a run with no hyperlink
     When I set the run hyperlink tooltip to "Just a tip"
     Then the reopened run hyperlink tooltip is "Just a tip"
      And the reopened run hyperlink address is None

  Scenario: Override a hyperlink run text color
    Given a run with an external hyperlink
     When I set the run hyperlink color to C00000
     Then the reopened run hyperlink color is C00000

  Scenario: Author a run-macro click action on a shape
    Given a shape
     When I set the shape click action to run macro "Recalc"
     Then the reopened shape click action is RUN_MACRO

  Scenario: Author a play-sound click action on a shape
    Given a shape
     When I attach a click sound to the shape
     Then the reopened shape click action has an embedded sound

  Scenario: Set an independent hover action on a shape
    Given a shape
     When I set the shape hover action address to "https://hover.example"
     Then the reopened shape hover action address is "https://hover.example"

  Scenario: Make a text run jump to another slide
    Given a run with no hyperlink
     When I make the run jump to a new slide
     Then the reopened run click action is NAMED_SLIDE

  Scenario: Add a hyperlink to a picture
    Given a picture on a slide
     When I set the picture hyperlink address to "https://pic.example"
     Then the reopened picture hyperlink address is "https://pic.example"
