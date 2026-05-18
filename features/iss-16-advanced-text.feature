Feature: Issue #16 — advanced text, auto-fit & internationalization
  In order to author rich, international, well-fitted text
  As a developer using python-pptx-extended
  I need run typography, CJK/complex-script fonts, columns, vertical
  text, RTL paragraphs, overflow detection and crash-free auto-fit

  Scenario: Author and read back superscript
    Given a blank slide text frame with one run
     When I set the run superscript
     Then the run reports superscript true

  Scenario: Author and read back double strikethrough
    Given a blank slide text frame with one run
     When I set the run strike to double
     Then the run reports strike double after round-trip

  Scenario: Author and read back a yellow highlight
    Given a blank slide text frame with one run
     When I set the run highlight to FFFF00
     Then the run reports highlight FFFF00 after round-trip

  Scenario: Author and read back character spacing
    Given a blank slide text frame with one run
     When I set the run character spacing to 2 points
     Then the run reports character spacing 2 points

  Scenario: East-Asian font set leaves Latin untouched
    Given a blank slide text frame with one run
     When I set east_asian to MS Gothic and name to Calibri
     Then latin is Calibri and east_asian is MS Gothic and they are independent

  Scenario: Two-column text box
    Given a blank slide text frame with one run
     When I set the text frame to 2 columns spaced 36 points
     Then the text frame reports 2 columns after round-trip

  Scenario: Vertical text direction
    Given a blank slide text frame with one run
     When I set the text direction to east asian vertical
     Then the text frame reports east asian vertical after round-trip

  Scenario: Arabic right-to-left paragraph
    Given a blank slide text frame with one run
     When I set the paragraph to Arabic right-to-left
     Then the paragraph reports rtl true after round-trip

  Scenario: Overflow detection flags oversized content
    Given a tiny text frame stuffed with text
     Then will_overflow reports true

  Scenario: fit_text survives a single unbreakable long word
    Given a tiny text frame with one very long word
     When I call fit_text on it
     Then no error is raised and auto_size is set

  Scenario: shrink_text_to_fit eagerly reduces the font scale
    Given a tiny text frame stuffed with text
     When I call shrink_text_to_fit
     Then the normAutofit fontScale is below 100
