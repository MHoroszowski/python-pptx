Feature: Slide duplicate — Slides.duplicate, Slide.duplicate
  In order to programmatically clone slides without lxml hacks
  As a developer using python-pptx
  I need a single API call that duplicates a slide and inserts it at a chosen position


  Scenario: Slides.duplicate(slide) inserts the copy after the source
    Given a Slides object containing 3 slides
     When I call slides.duplicate(slides[0])
     Then len(slides) is 4
      And the duplicate is at index 1
      And the source slide is still at index 0


  Scenario: Slides.duplicate(slide, index=N) inserts the copy at index N
    Given a Slides object containing 3 slides
     When I call slides.duplicate(slides[0], index=3)
     Then len(slides) is 4
      And the duplicate is at index 3


  Scenario: Slide.duplicate() returns a Slide with a new unique slide_id
    Given a Slides object containing 3 slides
     When I call slides[1].duplicate()
     Then len(slides) is 4
      And the duplicate slide_id is unique


  Scenario: Slides.duplicate raises IndexError for an out-of-range index
    Given a Slides object containing 3 slides
     Then calling slides.duplicate(slides[0], index=99) raises IndexError
