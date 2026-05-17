Feature: Add a slide layout to a slide master
  In order to build presentation templates programmatically
  As a developer using python-pptx
  I need to create new slide layouts on a slide master (issue #19 SF3)


  Scenario: SlideLayouts.add_layout() with no name
    Given a default presentation
     When I call slide_layouts.add_layout()
     Then the slide-master layout count increased by exactly 1
      And the new layout has a non-empty name


  Scenario: SlideLayouts.add_layout(name) sets the name
    Given a default presentation
     When I call slide_layouts.add_layout("Acceptance Layout")
     Then slide_layouts.get_by_name("Acceptance Layout") is the new layout


  Scenario: A presentation survives reopen after add_layout
    Given a default presentation
     When I call slide_layouts.add_layout("Persisted Layout")
      And I save and reopen the presentation
     Then the reopened presentation has a layout named "Persisted Layout"


  Scenario: A new layout is usable as the basis for a slide
    Given a default presentation
     When I call slide_layouts.add_layout("Slide Basis")
      And I add a slide based on the new layout
     Then the slide count increased by exactly 1


  Scenario: Presentation.save_as_potx writes a template content-type (SF2)
    Given a default presentation
     When I save the presentation as a potx
     Then the saved potx declares the template content-type
      And the in-memory presentation content-type is unchanged
      And the saved potx reopens as a valid presentation


  Scenario: Authoring a textbox directly on a slide master (SF5)
    Given a default presentation
     When I add a textbox to the slide master
      And I save and reopen the presentation
     Then the reopened slide master has the master textbox text


  Scenario: Duplicating a layout with copy_from (SF4)
    Given a default presentation
     When I call slide_layouts.add_layout("Copy Origin")
      And I add a textbox to the new layout
      And I copy the new layout with copy_from
     Then the copied layout has the same shape count as its source
      And the source layout is unchanged after copy_from


  Scenario: Applying a different layout to a slide (SF7)
    Given a default presentation
     When I add a slide on the default layout
      And I call slide_layouts.add_layout("Reassigned Layout")
      And I apply the new layout to that slide
      And I save and reopen the presentation
     Then the reopened slide uses the layout named "Reassigned Layout"
      And the reopened slide still resolves its slide master


  Scenario: Inserting a chart into a chart placeholder (SF8)
    Given a default presentation
     When I add a layout with a chart placeholder
      And I add a slide on that chart-placeholder layout
      And I insert a chart into the slide's chart placeholder
      And I save and reopen the presentation
     Then the reopened slide has exactly one chart
