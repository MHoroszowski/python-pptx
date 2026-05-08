Feature: Slide sections — read/write `<p14:sectionLst>`
  In order to organize slides into named groups in PowerPoint's slide pane
  As a developer using python-pptx
  I need to add, name, populate, and remove sections, with membership
  surviving slide reorder and removal


  Scenario: Presentation.sections is empty by default
    Given a Slides object containing 3 slides
     Then len(prs.sections) is 0


  Scenario: Add a section to a presentation
    Given a Slides object containing 3 slides
     When I call prs.sections.add_section("Intro")
     Then len(prs.sections) is 1
      And prs.sections[0].name is "Intro"


  Scenario: Add a slide to a section
    Given a Slides object containing 3 slides
     When I call prs.sections.add_section("Intro")
      And I call section.add_slide(prs.slides[0])
     Then len(section.slides) is 1


  Scenario: Slide membership survives a slide move
    Given a Slides object containing 3 slides
     When I call prs.sections.add_section("Body")
      And I call section.add_slide(prs.slides[0])
      And I call slides.move(slides[0], 2)
     Then section.slides still contains the moved slide
      And the moved slide is at presentation index 2


  Scenario: Remove a section cleans up extLst when last
    Given a Slides object containing 3 slides
     When I call prs.sections.add_section("Lonely")
      And I call prs.sections.remove(section)
     Then len(prs.sections) is 0
      And prs._element.extLst is None
