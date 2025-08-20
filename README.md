# Task & Idea Manager - Google Apps Script

A comprehensive task and idea management system built with Google Apps Script, Google Sheets, and a responsive HTML dashboard.

## Features

- **Template Creation**: Create standardized Idea and Task templates
- **Master Sheet Tracking**: Centralized tracking of all created templates
- **Interactive Dashboard**: Visual dashboard with charts and progress tracking
- **Automated Metrics**: Real-time calculation of task completion, overdue items, and progress
- **Sheet Management**: Automatic creation and sharing of template copies
- **Formula Integration**: IMPORTRANGE formulas for live data synchronization

## Project Structure

```
├── src/
│   └── gas/
│       ├── Code.js          # Main template creation and dashboard functions
│       ├── Code.gs          # Enhanced sheet management with IMPORTRANGE
│       └── dashboard.html   # Responsive HTML dashboard with charts
├── docs/                    # Documentation files
├── examples/               # Example configurations and use cases
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
   - `Code.js` or `Code.gs` (choose one based on your needs)
   - `dashboard.html`

### Setup

1. **Choose Your Implementation**:
   - **Code.js**: Basic implementation with direct template creation
   - **Code.gs**: Advanced implementation with IMPORTRANGE formulas and template management

2. **Create Required Sheets** (for Code.gs):
   - `Sheets_Master`: Main tracking sheet (auto-created)
   - `Template_List`: List of available templates with URLs

3. **Deploy Dashboard** (Optional):
   - For web app access: Deploy as Web App
   - For sidebar access: Use the menu items

## Usage

### Creating Templates

**Using Code.js:**
1. Open your Google Sheet
2. Use menu: `Task & Idea Manager` → `Create Idea Template` or `Create Task Template`
3. Templates are automatically created and tracked in `Sheets_Master`

**Using Code.gs:**
1. Set up `Template_List` sheet with template names and URLs
2. Add entries to `Sheets_Master` with template names and user information
3. Use menu: `Task Manager` → `Generate Sheets`

### Dashboard Access

1. **Modal Dialog**: Menu → `Open Dashboard`
2. **Sidebar**: Menu → `Open Dashboard (Sidebar)` (Code.gs)
3. **Web App**: Deploy as web app for standalone access

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

### Code.js Functions
- `createIdeaTemplate()`: Creates new idea tracking template
- `createTaskTemplate()`: Creates new task tracking template
- `openDashboard()`: Opens interactive dashboard
- `getMasterData()`: Retrieves data for dashboard display

### Code.gs Functions
- `processSheetsMaster()`: Creates sheet copies from templates
- `updateMetrics()`: Updates/ensures IMPORTRANGE formulas
- `showDashboardSidebar()`: Opens dashboard in sidebar
- `detectTasksSheetName()`: Auto-detects task sheet names

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

---

**Note**: This system requires Google Workspace access and appropriate permissions for Google Sheets, Drive, and Apps Script.