Feature: Table merge robustness — range-style merge_cells / split_cells
  In order to assemble tables with block merges programmatically
  As a developer using python-pptx
  I need range-style idempotent Table.merge_cells / Table.split_cells, and read-only inspection of gridSpan/rowSpan/hMerge/vMerge


  Scenario: Range merge a 2x3 block
    Given a 3x3 table on a fresh slide
     When I call table.merge_cells row=(0,1) col=(0,2)
     Then cell (0,0) has gridSpan=3 and rowSpan=2
      And cell (0,0) is_merge_origin is True
      And cell (1,1) hMerge is True
      And cell (1,1) vMerge is True


  Scenario: Range merge is idempotent on exact re-merge
    Given a 3x3 table on a fresh slide
     When I call table.merge_cells row=(0,1) col=(0,2)
      And I call table.merge_cells row=(0,1) col=(0,2)
     Then cell (0,0) has gridSpan=3 and rowSpan=2


  Scenario: Range merge accepts Python range objects
    Given a 3x3 table on a fresh slide
     When I call table.merge_cells with range(0,2) and range(0,3)
     Then cell (0,0) has gridSpan=3 and rowSpan=2


  Scenario: Range merge raises on partial overlap
    Given a 3x3 table on a fresh slide
     When I call table.merge_cells row=(0,0) col=(0,1)
     Then calling table.merge_cells row=(0,1) col=(0,1) raises ValueError


  Scenario: Single-cell range merge is a no-op
    Given a 3x3 table on a fresh slide
     When I call table.merge_cells row=(0,0) col=(0,0)
     Then cell (0,0) has gridSpan=1 and rowSpan=1


  Scenario: Range split unmerges a block
    Given a 3x3 table on a fresh slide
     When I call table.merge_cells row=(0,1) col=(0,2)
      And I call table.split_cells row=(0,1) col=(0,2)
     Then cell (0,0) has gridSpan=1 and rowSpan=1
      And cell (1,1) hMerge is False
      And cell (1,1) vMerge is False


  Scenario: Range split is idempotent on un-merged ranges
    Given a 3x3 table on a fresh slide
     When I call table.split_cells row=(0,2) col=(0,2)
     Then cell (0,0) has gridSpan=1 and rowSpan=1


  Scenario: Range split raises when merge crosses range boundary
    Given a 3x3 table on a fresh slide
     When I call table.merge_cells row=(0,1) col=(0,2)
     Then calling table.split_cells row=(0,0) col=(0,1) raises ValueError
