Feature: Table sizing & ergonomics — row_count / column_count / dimensions + sizing round-trip
  In order to inspect a table's shape and rely on persistent row heights and column widths
  As a developer using python-pptx
  I need read-only count properties on Table and round-trip preservation of explicit sizes


  Scenario: row_count returns the number of rows
    Given a 3x4 table on a fresh slide
     Then table.row_count is 3


  Scenario: column_count returns the number of columns
    Given a 3x4 table on a fresh slide
     Then table.column_count is 4


  Scenario: dimensions returns a (rows, cols) tuple
    Given a 3x4 table on a fresh slide
     Then table.dimensions is (3, 4)


  Scenario: row_count updates after rows.add()
    Given a 3x4 table on a fresh slide
     When I add a row to the table
     Then table.row_count is 4


  Scenario: column_count updates after columns.remove()
    Given a 3x4 table on a fresh slide
     When I remove column 0 from the table
     Then table.column_count is 3


  Scenario: Row height round-trips through save/reload
    Given a 3x4 table on a fresh slide
     When I set row 0 height to 500000 EMU
      And I save and reload the presentation via stream
     Then the reloaded row 0 height is 500000


  Scenario: Column width round-trips through save/reload
    Given a 3x4 table on a fresh slide
     When I set column 1 width to 1000000 EMU
      And I save and reload the presentation via stream
     Then the reloaded column 1 width is 1000000
