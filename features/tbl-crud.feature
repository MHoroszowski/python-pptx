Feature: Table row/column CRUD — add and remove
  In order to assemble tables programmatically without lxml hacks
  As a developer using python-pptx
  I need to add and remove rows and columns on existing tables


  Scenario: Append a row to a table
    Given a 2x3 table on a fresh slide
     When I call table.rows.add()
     Then the table has 3 rows
      And the table has 3 columns


  Scenario: Insert a row at index 1
    Given a 2x3 table on a fresh slide
     When I call table.rows.add(at=1)
     Then the table has 3 rows
      And the new row is at index 1


  Scenario: Append a column to a table
    Given a 2x3 table on a fresh slide
     When I call table.columns.add()
     Then the table has 4 columns
      And every row has 4 cells


  Scenario: Remove a row by index
    Given a 3x2 table on a fresh slide
     When I call table.rows.remove(1)
     Then the table has 2 rows


  Scenario: Remove a column by index
    Given a 2x3 table on a fresh slide
     When I call table.columns.remove(1)
     Then the table has 2 columns
      And every row has 2 cells


  Scenario: Removing a row with vertical merge raises ValueError
    Given a 3x2 table on a fresh slide with a vertical merge between rows 0 and 1
     Then calling table.rows.remove(0) raises ValueError


  Scenario: Removing a column with horizontal merge raises ValueError
    Given a 2x3 table on a fresh slide with a horizontal merge between columns 0 and 1
     Then calling table.columns.remove(0) raises ValueError
