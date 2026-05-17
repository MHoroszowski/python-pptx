Feature: Modern threaded comments on a slide (issue #25)
  In order to round-trip review feedback in a presentation
  As a developer using python-pptx
  I need to add, read, reply to, and remove threaded comments on a slide

  Scenario: Add a threaded comment and read it back after save/reopen
    Given a blank slide with no comments
    When I add a comment "Looks great" by author "Ada"
    Then the slide has 1 comment
    And after save and reopen the first comment text is "Looks great" by author "Ada"

  Scenario: Reply to a comment threads under it and round-trips
    Given a blank slide with no comments
    When I add a comment "Question?" by author "Ada"
    And I reply "Here is the answer" by author "Babbage"
    Then after save and reopen the first comment has 1 reply with text "Here is the answer"

  Scenario: Removing the last comment leaves a valid reopenable file
    Given a blank slide with no comments
    When I add a comment "Temporary" by author "Ada"
    And I remove that comment
    Then the slide has 0 comments
    And after save and reopen the slide has 0 comments

  Scenario: Resolving a threaded comment round-trips
    Given a blank slide with no comments
    When I add a comment "Please review" by author "Ada"
    And I resolve that comment
    Then after save and reopen the first comment is resolved

  Scenario: A legacy comment and a modern comment coexist on round-trip
    Given a slide carrying a pre-existing legacy comment
    When I add a modern comment "Modern note" by author "Mary"
    Then after save and reopen both the legacy and modern comments are present
    And after save and reopen the legacy author ids are intact
