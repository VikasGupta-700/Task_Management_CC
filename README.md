# Task Management System - Google Apps Script

An enhanced task management system built with Google Apps Script, Google Sheets, and a responsive HTML dashboard.

## Features

- **Bulk Sheet Creation**: Create multiple task sheets from master template list
- **Single Sheet Creation**: Create individual sheets from task master
- **Interactive Dashboard**: Visual dashboard with charts and progress tracking
- **Automated Metrics**: Real-time calculation of task completion, overdue items, and progress
- **Smart Sheet Management**: Automatic creation, sharing, and status tracking
- **IMPORTRANGE Integration**: Live data synchronization with comprehensive error handling

## Project Structure

```
├── src/
│   └── gas/
│       ├── Code.gs          # Enhanced task management system with bulk creation
│       └── dashboard.html   # Responsive HTML dashboard with charts
├── test-data*.csv          # Test data files for various scenarios
├── TEST-SCENARIOS.md       # Test scenarios documentation
└── README.md              # This file
```

## Getting Started

### Prerequisites

- Google Apps Script access
- Google Sheets
- Google Drive permissions for file creation and sharing

### Installation

1. **Clone or Download**: Get the project files to your local machine
2. **Create New GAS Project**: Go to [script.google.com](https://script.google.com) and create a new project
3. **Upload Files**: Copy the contents from `src/gas/` to your GAS project:
   - `Code.gs` - Main enhanced task management system
   - `dashboard.html` - Interactive dashboard

### Setup

1. **Create Required Sheets**:
   - `Sheets_Master`: Main tracking sheet with columns for template tracking
   - `Template_List`: List of available templates with their URLs

2. **Deploy Dashboard** (Optional):
   - For web app access: Deploy as Web App with execute permissions
   - Access via menu: `Task Management System` → `Open Dashboard`

## Usage

### Creating Task Sheets

1. **Set up Template List**: Create `Template_List` sheet with template names and URLs
2. **Add to Sheets_Master**: Add entries with template names and user information
3. **Bulk Creation**: Use menu `Task Management System` → `Create Sheets`
4. **Individual Creation**: Process single rows where Column E (status) is empty

### Dashboard Access

1. **Modal Dialog**: Menu → `Task Management System` → `Open Dashboard`
2. **Web App**: Deploy as web app for standalone dashboard access
3. **Direct Function**: Call `getDashboardData()` for programmatic access

### Sheet Structure

#### Idea/Task Templates
| Column | Description |
|--------|-------------|
| A | Sr. No |
| B | Title |
| C | Description |
| D | Date (Idea Date / Allocated Date) |
| E | Planned Implementation/Completion Date |
| F | Actual Implementation/Completion Date |
| G | Status |
| H | Remarks / Issues |
| I | On Time / Delayed (Formula) |
| J | Delay Days (Formula) |
| K | Estimated? (Formula) |
| L | Days Since Allocated (Formula) |

#### Sheets_Master
| Column | Description |
|--------|-------------|
| A | Sheet Template |
| B | Shared With |
| C | Email ID |
| D | Sheet URL |
| E | Sheet Status |
| F | Date Created |
| G | Total tasks (Formula) |
| H | Completed tasks (Formula) |
| I | Pending tasks (Formula) |
| J | Overdue tasks (Formula) |
| K | Progress % (Formula) |

## Key Functions

### Core Functions
- `createSheets()`: Bulk creates sheets for all pending rows in Sheets_Master
- `processSheetCreation()`: Creates individual sheet from template
- `openDashboard()`: Opens interactive dashboard modal
- `getDashboardData()`: Retrieves comprehensive dashboard data
- `updateTaskMasterRow()`: Updates master sheet with new sheet information
- `insertImportRangeFormulas()`: Adds live data synchronization formulas
- `getOverdueTaskDetails()`: Calculates overdue task metrics

## Dashboard Features

- **Summary Cards**: Overview of templates, tasks, completed, and overdue items
- **Template Cards**: Individual cards showing progress and metrics for each template
- **Progress Charts**: Bar chart showing completion progress
- **Status Distribution**: Doughnut chart showing task status breakdown
- **Responsive Design**: Works on desktop and mobile devices

## Customization

### Status Values
The system recognizes these status values as "completed":
- "Completed"
- "Done" 
- "Closed"
- "Complete"

### Formula Customization
Modify the formulas in `setFormulasForRow()` function to match your specific needs:
- Total tasks calculation
- Completed tasks criteria
- Overdue logic
- Progress calculation

### Dashboard Styling
Edit `dashboard.html` to customize:
- Color schemes
- Layout and spacing
- Chart types and configurations
- Card designs

## Troubleshooting

### Common Issues

1. **Permission Errors**: Ensure proper Google Drive and Sheets permissions
2. **IMPORTRANGE Issues**: May require manual approval on first use
3. **Template Not Found**: Verify Template_List sheet setup (Code.gs)
4. **Dashboard Empty**: Check console logs and verify getMasterData() function

### Debugging

- Enable console logging in dashboard.html
- Check Google Apps Script execution logs
- Verify sheet names and ranges match code expectations

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is open source. Feel free to use, modify, and distribute according to your needs.

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review Google Apps Script documentation
3. Create an issue in the repository

## Version History

- **v1.0**: Initial release with basic template creation
- **v1.1**: Added IMPORTRANGE formulas and enhanced tracking  
- **v1.2**: Responsive dashboard with charts and improved UI
- **v2.0**: Complete system overhaul with bulk creation and enhanced error handling
- **v2.1**: Menu simplification, URL validation, and dashboard improvements

---

**Note**: This system requires Google Workspace access and appropriate permissions for Google Sheets, Drive, and Apps Script.