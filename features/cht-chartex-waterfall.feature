Feature: ChartEx waterfall charts
  In order to add Office 2016 waterfall charts to a slide
  As a developer using python-pptx
  I need the ChartEx writer path to create modern chart graphic frames


  Scenario Outline: Add a waterfall chart through either public entry point
    Given a blank slide
      And ChartEx waterfall data case q4-total
     When I add the ChartEx waterfall via <add-path>
     Then the active ChartEx frame exposes ChartEx but not a classic chart
      And the active ChartEx chart type is waterfall
      And the active ChartEx series is named Revenue
      And the active ChartEx series values are 100.0, 50.0, -30.0, 80.0, 200.0

    Examples: public waterfall entry points
      | add-path    |
      | add_chart   |
      | add_chartex |


  Scenario Outline: Waterfall category labels are preserved on creation
    Given a blank slide
      And ChartEx waterfall data case <data-case>
     When I add the ChartEx waterfall via <add-path>
     Then the active ChartEx category labels are <category-labels>

    Examples: category label sets
      | data-case       | add-path    | category-labels                      |
      | q4-total        | add_chart   | Q1, Q2, Q3, Q4, Total               |
      | q4-total        | add_chartex | Q1, Q2, Q3, Q4, Total               |
      | cash-bridge     | add_chart   | Start, Sales, Returns, Ops, Tax, End |
      | cash-bridge     | add_chartex | Start, Sales, Returns, Ops, Tax, End |
      | regional-rollup | add_chart   | East, West, Midwest, Online, Total  |
      | regional-rollup | add_chartex | East, West, Midwest, Online, Total  |


  Scenario Outline: Waterfall subtotal markers survive round-trip
    Given a blank slide
      And ChartEx waterfall data case <data-case>
     When I add the ChartEx waterfall via <add-path>
      And I round-trip the presentation for ChartEx inspection
     Then the active ChartEx frame exposes ChartEx but not a classic chart
      And the active ChartEx subtotal indices are <subtotal-indices>
      And the active ChartEx category labels are <category-labels>

    Examples: subtotal preservation cases
      | data-case       | add-path    | subtotal-indices | category-labels                      |
      | q4-total        | add_chart   | 4                | Q1, Q2, Q3, Q4, Total               |
      | q4-total        | add_chartex | 4                | Q1, Q2, Q3, Q4, Total               |
      | cash-bridge     | add_chart   | 0, 5             | Start, Sales, Returns, Ops, Tax, End |
      | cash-bridge     | add_chartex | 0, 5             | Start, Sales, Returns, Ops, Tax, End |
      | regional-rollup | add_chart   | 4                | East, West, Midwest, Online, Total  |
      | regional-rollup | add_chartex | 4                | East, West, Midwest, Online, Total  |
