Feature: Accessibility Phase B — reading order, lint helper, decorative alias
  In order to ship Section 508 / WCAG / ADA compliant decks
  As a developer using python-pptx
  I need to query reading order, identify shapes lacking alt text,
  and toggle the "hidden from accessibility" flag


  Scenario: is_hidden_from_accessibility mirrors is_decorative
    Given a slide with one textbox
     When I set shape.is_hidden_from_accessibility to True
     Then shape.is_decorative is True


  Scenario: reading_order returns shapes in z-order
    Given a slide with three textboxes labelled A, B, C
     Then reading_order produces shapes in the order A, B, C


  Scenario: reading_order setter reorders the slide's spTree
    Given a slide with three textboxes labelled A, B, C
     When I set reading_order to (C, A, B)
     Then iteration produces shapes in the order C, A, B


  Scenario: accessibility_issues flags shapes without alt text or decorative flag
    Given a slide with three textboxes labelled A, B, C
     When I tag A with alt text and B as decorative
     Then accessibility_issues returns just C
