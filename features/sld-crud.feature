Feature: Slide CRUD — remove, move, indexed add
  In order to programmatically assemble decks without lxml hacks
  As a developer using python-pptx
  I need to add slides at a chosen position, reorder them, and remove them


  Scenario: Slides.add_slide(slide_layout, index=0) inserts at the head
    Given a Slides object containing 3 slides
     When I call slides.add_slide(slide_layout, index=0)
     Then len(slides) is 4
      And the new slide is at index 0


  Scenario: Slides.add_slide(slide_layout, index=2) inserts in the middle
    Given a Slides object containing 3 slides
     When I call slides.add_slide(slide_layout, index=2)
     Then len(slides) is 4
      And the new slide is at index 2


  Scenario: Slides.move(slide, new_index) reorders slides
    Given a Slides object containing 3 slides
     When I call slides.move(slides[0], 2)
     Then len(slides) is 3
      And the slide order matches the original [1, 2, 0]


  Scenario: Slides.remove(slide) drops a slide and its rel
    Given a Slides object containing 3 slides
     When I call slides.remove(slides[1])
     Then len(slides) is 2
      And the surviving slide order matches the original [0, 2]


  Scenario: Slide.delete() removes the slide from its presentation
    Given a Slides object containing 3 slides
     When I call slides[1].delete()
     Then len(slides) is 2
      And the surviving slide order matches the original [0, 2]
