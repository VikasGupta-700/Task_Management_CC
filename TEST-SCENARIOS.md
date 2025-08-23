# Test Scenarios for Task & Idea Manager

## Overview
This document outlines comprehensive test scenarios to validate all functionality of the Task & Idea Manager system.

## Test Data Files
- `test-data.csv` - Standard task test data with realistic scenarios (25 tasks)
- `test-data-ideas.csv` - Standard idea test data with realistic scenarios (25 ideas)
- `test-data-edge-cases.csv` - Edge cases and boundary conditions (30 tasks)

## Testing Instructions

### 1. Import Test Data
1. Create a new Google Sheet
2. Import one of the test data files:
   - `test-data.csv` - For standard task testing
   - `test-data-ideas.csv` - For idea template testing  
   - `test-data-edge-cases.csv` - For edge case testing
3. Ensure column headers match the correct format:
   
   **For Task Sheets:**
   - Sr. No, Task Title, Task Description, Allocated Date, Planned Completion Date, Actual Completion Date, Status
   
   **For Idea Sheets:**
   - Sr. No, Idea Title, Idea Description, Idea Date, Planned Implementation Date, Actual Implementation Date, Status

4. Use the Task Manager to create a new template pointing to this sheet

### 2. Total Tasks Calculation Tests

#### Scenario A: Basic Count Verification
- **Expected**: Total tasks should equal number of data rows (excluding header)
- **Test Data**: Use `test-data.csv` (25 tasks)
- **Formula**: `SUMPRODUCT(--(LEN(IMPORTRANGE(URL,"B2:B1000"))>0))` - counts non-empty Task Titles from row 2 onwards
- **Verification**: Column G should show 25

#### Scenario B: Empty Sheet Test (New Template)
- **Setup**: Newly created sheet with headers only, no data rows
- **Expected**: Total tasks = 0
- **Issue Fixed**: Previous formula was counting header row or empty cells with spaces
- **New Formula**: `SUMPRODUCT(--(LEN(IMPORTRANGE(URL,"B2:B1000"))>0))` only counts cells with actual content
- **Verification**: Newly created sheets should show 0 total tasks, not 1

#### Scenario C: Mixed Data Test
- **Setup**: Sheet with some empty rows between data
- **Expected**: Count only non-empty task titles (column B)
- **Verification**: COUNTA should ignore empty cells in B2:B range

### 3. Status-Based Counting Tests

#### Completed Tasks (Column H)
- **Formula**: `SUM(COUNTIF(IMPORTRANGE(URL,"G:G"),{"Completed","Done","Closed","Complete"}))`
- **Test Cases**:
  - Standard "Completed" status (column G)
  - Alternative completion statuses: "Done", "Closed", "Complete"
  - Case sensitivity validation
- **Expected Results**:
  - `test-data.csv`: 4 completed tasks (rows 1, 2, 7, 16)
  - `test-data-edge-cases.csv`: 8 completed tasks

#### Pending Tasks (Column I)
- **Formula**: `G${rowIndex}-H${rowIndex}` (Total - Completed)
- **Test Cases**:
  - "Pending" status tasks
  - "In Progress" status tasks
  - Any non-completed status
- **Expected**: Should dynamically calculate: Total tasks minus completed tasks

#### Overdue Tasks (Column J)
- **Formula**: `COUNTIFS(IMPORTRANGE(URL,"E:E"),"<"&TODAY(),IMPORTRANGE(URL,"G:G"),"<>Completed")`
- **Criteria**: Planned Completion Date (column E) < TODAY() AND Status (column G) != "Completed"
- **Test Cases**:
  - Tasks with past planned dates
  - Tasks without planned dates (should not count as overdue)
  - Weekend and holiday dates
- **Expected Results**:
  - `test-data.csv`: Tasks with past planned completion dates that aren't completed
  - `test-data-edge-cases.csv`: Row 4 (Past due date test)

### 4. Date Handling Tests

#### Date Format Variations
- **Test Cases**:
  - MM/DD/YYYY format
  - DD/MM/YYYY format  
  - YYYY-MM-DD format
  - Text dates ("January 1, 2024")
- **Expected**: All formats should be properly parsed

#### Special Dates
- **Test Cases**:
  - Leap year dates (2024-02-29)
  - Year boundaries (2024-12-31, 2025-01-01)
  - Weekend dates
  - Holiday dates
- **Expected**: All dates should be handled correctly

#### Null/Empty Dates
- **Test Cases**:
  - Empty planned date cells
  - Empty actual date cells
- **Expected**: Should not cause errors, handle gracefully

### 5. Dashboard Integration Tests

#### Card Display
- **Test**: Create multiple templates with test data
- **Expected**: Dashboard should show 10 cards simultaneously on laptop screen
- **Verification**: Cards should be properly sized and responsive

#### Progress Calculation
- **Formula**: Completed / Total * 100
- **Test Cases**:
  - 0% progress (no completed tasks)
  - 100% progress (all tasks completed)
  - Partial progress (mixed statuses)
- **Expected**: Accurate percentage calculations

#### Summary Statistics
- **Test**: Aggregate data across multiple templates
- **Verification**:
  - Total templates count
  - Sum of all tasks across templates
  - Sum of completed tasks
  - Sum of overdue tasks

### 6. Formula Testing

#### IMPORTRANGE Function
- **Test Cases**:
  - Valid Google Sheets URLs
  - Invalid URLs (should show error gracefully)
  - Sheets with different column arrangements
  - Private sheets (permission testing)

#### Error Handling
- **Test Cases**:
  - IFERROR wrapper functionality
  - Division by zero scenarios
  - Missing data scenarios
- **Expected**: Formulas should never show #ERROR, always fallback to 0 or appropriate default

### 7. Performance Tests

#### Large Data Sets
- **Test Cases**:
  - 100 tasks
  - 500 tasks  
  - 1000+ tasks
- **Expected**: Reasonable performance, no timeouts

#### Multiple Templates
- **Test Cases**:
  - 5 templates
  - 10 templates
  - 20+ templates
- **Expected**: Dashboard loads within acceptable time

### 8. Edge Case Validation

#### Special Characters
- **Test Cases**:
  - Unicode characters
  - Emojis in task titles
  - Special punctuation
  - HTML/script tags (security)
- **Expected**: All characters display correctly, no security issues

#### Data Validation
- **Test Cases**:
  - Very long task titles (500+ characters)
  - Very long descriptions (1000+ characters)
  - Missing required fields
- **Expected**: Graceful handling without breaking layout

#### Dependency Testing
- **Test Cases**:
  - Valid dependencies
  - Circular dependencies
  - Non-existent dependencies
  - Multiple dependencies
- **Expected**: Dependencies tracked correctly, circular deps detected

## Validation Checklist

### Pre-Testing Setup
- [ ] Test data imported correctly
- [ ] Column headers match expected format
- [ ] Formulas applied to correct ranges
- [ ] IMPORTRANGE permissions granted

### Core Functionality Tests
- [ ] Total tasks count excludes header row
- [ ] Completed tasks count all completion statuses
- [ ] Overdue calculation works correctly
- [ ] Progress percentages calculate accurately
- [ ] Date parsing handles various formats
- [ ] Empty data cells handled gracefully

### Dashboard Tests
- [ ] Cards display properly sized (10 visible on laptop)
- [ ] Summary statistics aggregate correctly
- [ ] Charts render with accurate data
- [ ] Responsive design works on different screen sizes

### Error Handling Tests
- [ ] Invalid URLs handled gracefully
- [ ] Missing permissions show appropriate message
- [ ] Formula errors don't break display
- [ ] Malformed data doesn't cause crashes

### Performance Tests
- [ ] Large datasets load within 30 seconds
- [ ] Dashboard renders within 10 seconds
- [ ] Multiple templates don't cause timeouts
- [ ] Memory usage remains reasonable

## Expected Results Summary

### test-data.csv (25 tasks)
- Total Tasks: 25
- Completed: 4 (Completed status)
- Pending: 21 (Total - Completed)
- Overdue: Depends on current date vs planned completion dates

### test-data-ideas.csv (25 ideas)
- Total Ideas: 25
- Implemented: 4 (Implemented status)
- Under Review: 11 
- Approved: 10
- Overdue: Depends on current date vs planned implementation dates

### test-data-edge-cases.csv (30 tasks)
- Total Tasks: 30
- Completed: 8 (Completed, Done, Closed, Complete statuses)
- Pending: 22 (Total - Completed)  
- Overdue: 1 (Row 4 - Past due date test)

## Troubleshooting Guide

### Common Issues
1. **Formula shows #ERROR**: Check IMPORTRANGE permissions
2. **Wrong task count**: Verify B2:B range instead of A:A or B:B
3. **Dates not calculating**: Check date format consistency
4. **Dashboard not loading**: Verify all sheet URLs are accessible
5. **Cards too large**: Confirm CSS changes applied correctly

### Debug Steps
1. Check browser console for JavaScript errors
2. Verify Google Sheets permissions
3. Test formulas individually in Google Sheets
4. Validate CSV import formatting
5. Check network connectivity for IMPORTRANGE