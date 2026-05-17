Feature: ChartEx round-trip behavior
  In order to preserve modern chart parts in saved presentations
  As a developer using python-pptx
  I need ChartEx content and classic-chart content to round-trip correctly


  Scenario Outline: A waterfall ChartEx part survives save and reopen
    Given a blank slide
      And ChartEx waterfall data case q4-total
     When I add the ChartEx waterfall via <add-path>
      And I round-trip the presentation for ChartEx inspection
     Then the active ChartEx frame exposes ChartEx but not a classic chart
      And the active ChartEx part content type is the ChartEx content type
      And the active ChartEx partname contains chartEx
      And the active ChartEx series values are 100.0, 50.0, -30.0, 80.0, 200.0

    Examples: round-trip entry points
      | add-path    |
      | add_chart   |
      | add_chartex |


  Scenario: Classic and ChartEx chart frames coexist on one slide across round-trip
    Given a blank slide
      And ChartEx waterfall data case regional-rollup
     When I add a classic chart beside a ChartEx waterfall
     Then the slide has one classic chart frame and one ChartEx frame
      And the classic chart frame still exposes only a classic chart
      And the ChartEx frame still exposes only a ChartEx chart
     When I round-trip the presentation for ChartEx inspection
     Then the slide has one classic chart frame and one ChartEx frame
      And the classic chart frame still exposes only a classic chart
      And the ChartEx frame still exposes only a ChartEx chart
      And the active ChartEx series values are 30.0, -10.0, 25.0, 15.0, 60.0


  Scenario: A blank presentation exposes no ChartEx frames
    Given a blank slide
     Then the slide has no ChartEx graphic frames


  Scenario: A blank presentation saves without any ChartEx package parts
    Given a blank slide
     When I round-trip the presentation for ChartEx inspection
     Then the slide has no ChartEx graphic frames
      And the saved package contains no ChartEx partnames
      And the saved package contains no ChartEx content type declaration
