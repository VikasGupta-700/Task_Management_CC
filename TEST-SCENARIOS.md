# Test Scenarios for Task & Idea Manager

## Overview
This document outlines comprehensive test scenarios to validate all functionality of the Task & Idea Manager system.

## Test Data Files
- `test-data.csv` - Standard test data with realistic scenarios
- `test-data-edge-cases.csv` - Edge cases and boundary conditions

## Testing Instructions

### 1. Import Test Data
1. Create a new Google Sheet
2. Import either `test-data.csv` or `test-data-edge-cases.csv`
3. Ensure column headers match: ID, Task Title, Task Description, Task Type, Planned Date, Actual Date, Status, Priority, Health, Task Dependencies, Confirmed
4. Use the Task Manager to create a new template pointing to this sheet

### 2. Total Tasks Calculation Tests

#### Scenario A: Basic Count Verification
- **Expected**: Total tasks should equal number of data rows (excluding header)
- **Test Data**: Use `test-data.csv` (25 tasks)
- **Verification**: Column G should show 25, not 26

#### Scenario B: Empty Sheet Test
- **Setup**: Create sheet with headers only, no data rows
- **Expected**: Total tasks = 0
- **Verification**: Formula should handle empty ranges gracefully

#### Scenario C: Mixed Data Test
- **Setup**: Sheet with some empty rows between data
- **Expected**: Count only non-empty task titles
- **Verification**: COUNTA should ignore empty cells in B2:B range

### 3. Status-Based Counting Tests

#### Completed Tasks (Column H)
- **Test Cases**:
  - Standard "Completed" status
  - Alternative completion statuses: "Done", "Closed", "Complete"
  - Case sensitivity validation
- **Expected Results**:
  - `test-data.csv`: 4 completed tasks (IDs: 1, 2, 7, 16)
  - `test-data-edge-cases.csv`: 6 completed tasks

#### Pending Tasks (Column I)
- **Calculation**: Total - Completed - Overdue
- **Test Cases**:
  - "Pending" status tasks
  - "In Progress" status tasks
- **Expected**: Should dynamically calculate based on other columns

#### Overdue Tasks (Column J)
- **Criteria**: Planned Date < TODAY() AND Status != "Completed"
- **Test Cases**:
  - Tasks with past planned dates
  - Tasks without planned dates (should not count as overdue)
  - Weekend and holiday dates
- **Expected Results**:
  - `test-data.csv`: 2 overdue tasks (IDs: 8, 15, 23)

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
- Completed: 4
- In Progress: 4  
- Pending: 15
- Overdue: 2

### test-data-edge-cases.csv (30 tasks)
- Total Tasks: 30
- Completed: 6
- In Progress: 4
- Pending: 18
- Overdue: 2

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