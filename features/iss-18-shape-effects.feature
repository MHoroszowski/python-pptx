Feature: Issue #18 — shape effects, 3-D, arrowheads, flip, duplicate
  In order to author rich PowerPoint shapes that open without repair
  As a developer using python-pptx-extended
  I need glow / reflection / soft-edge effects, preset 3-D, flip,
  shape duplication, and the issue-named arrowhead / connector API

  Scenario: Author and read back a glow color and radius
    Given a blank slide with one rectangle
     When I set the glow color to FF0000 and radius to 20pt
     Then the rectangle reports glow radius 20pt and color FF0000

  Scenario: Author and read back a reflection
    Given a blank slide with one rectangle
     When I set the reflection blur radius to 3pt and distance to 7pt
     Then the rectangle reports reflection blur radius 3pt

  Scenario: Author and read back a soft edge
    Given a blank slide with one rectangle
     When I set the soft edge radius to 5pt
     Then the rectangle reports soft edge radius 5pt

  Scenario: Author and read back a preset 3-D camera
    Given a blank slide with one rectangle
     When I set the 3-D camera preset to orthographicFront
     Then the rectangle reports camera preset orthographicFront
      And the scene has a light rig

  Scenario: Author and read back a 3-D extrusion
    Given a blank slide with one rectangle
     When I set the extrusion height to 18pt
     Then the rectangle reports extrusion height 18pt

  Scenario: Flip a shape vertically and read it back after round-trip
    Given a blank slide with one rectangle
     When I flip the rectangle vertically and round-trip the file
     Then the reopened rectangle is flipped vertically

  Scenario: Flip a shape horizontally
    Given a blank slide with one rectangle
     When I flip the rectangle horizontally
     Then the rectangle is flipped horizontally

  Scenario: Duplicate a shape produces two distinct shapes
    Given a blank slide with one rectangle
     When I duplicate the rectangle
     Then the slide has two rectangles with distinct shape ids

  Scenario: Arrow-ended connector via the head_end API round-trips
    Given a blank slide with one connector
     When I set the tail end arrowhead to a triangle and round-trip the file
     Then the reopened connector tail end is a triangle

  Scenario: Group child reports correct world-space coordinates
    Given a group scaled two-to-one containing one rectangle
     Then the rectangle slide_width is double its local width
