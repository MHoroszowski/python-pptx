Feature: ChartEx Phase-C writers and replace_data
  In order to author every Office-2016 modern chart type
  As a developer using python-pptx
  I need each ChartEx type to write, round-trip, and support replace_data


  Scenario Outline: Each ChartEx type writes and round-trips
    Given a blank slide
     When I add a ChartEx <member-name> chart
     Then the slide has a ChartEx graphic frame
      And the saved package contains a ChartEx part
      And the ChartEx round-trips preserving its part

    Examples: ChartEx writable types
      | member-name |
      | WATERFALL   |
      | TREEMAP     |
      | SUNBURST    |
      | FUNNEL      |
      | BOX_WHISKER |
      | HISTOGRAM   |
      | PARETO      |


  Scenario Outline: replace_data updates each ChartEx type and round-trips
    Given a blank slide
     When I add a ChartEx <member-name> chart
      And I replace the ChartEx <member-name> data with a smaller dataset
     Then the reopened ChartEx reflects the replaced data
      And the ChartEx round-trips preserving its part

    Examples: replace_data types
      | member-name |
      | WATERFALL   |
      | TREEMAP     |
      | SUNBURST    |
      | FUNNEL      |
      | HISTOGRAM   |
      | PARETO      |


  Scenario: replace_data rejects a chart-type mismatch
    Given a blank slide
     When I attempt to replace a FUNNEL ChartEx with HISTOGRAM data
     Then a chart-type mismatch error is raised
