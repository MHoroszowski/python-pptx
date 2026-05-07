Feature: Presentation.append_from — cross-deck slide copy
  In order to assemble a deck from slides scattered across multiple .pptx files
  As a developer using python-pptx
  I need a single API call that copies slides between two presentations,
  porting the slide-master / layout / theme and deduping images at the target


  Scenario: Presentation.append_from(other) appends every source slide
    Given two seeded presentations source and target
     When I call target.append_from(source)
     Then target's slide count grew by source's slide count
      And target's master count grew by 1


  Scenario: Presentation.append_from(other, slide_indexes) appends selected slides
    Given two seeded presentations source and target
     When I call target.append_from(source, slide_indexes=[0, 1])
     Then target's slide count grew by 2


  Scenario: append_from with empty slide_indexes is a no-op
    Given two seeded presentations source and target
     When I call target.append_from(source, slide_indexes=[])
     Then target's slide count is unchanged


  Scenario: append_from raises IndexError for out-of-range index
    Given two seeded presentations source and target
     Then calling target.append_from(source, slide_indexes=[99]) raises IndexError
