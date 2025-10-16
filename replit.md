# Contract Review Checklist System

## Overview
A Flask-based web application for managing contract review checklists with Excel-like interface, user authentication, and role-based access control.

## Features
- **User Authentication**: Login system with Admin and User roles
- **Default Credentials**:
  - Admin: username `admin`, password `admin123`
  - User: username `user`, password `user123`
- **Admin Panel**: Master data management for users (username, name, email, department, password, role)
- **Excel Integration**: Upload Excel files with multiple worksheets and display in interactive spreadsheet format
- **Excel Download**: Download current checklist data as Excel file with same template format - keeps raw Excel editable
- **Multi-Worksheet Support**: Navigate between all Excel worksheets via tabs (CR 1-6, PED-1, PED-2, Lead Time, etc.)
- **Editable Grid**: Handsontable-powered Excel-like interface with sorting, filtering, and inline editing
- **Auto-save**: Changes are automatically saved to SQLite database per worksheet
- **Contract Review Items**: Supports all 12 worksheets from the contract review checklist including Index Sheet, CR 1-6, PED-1, PED-2, Lead Time, and more

## Technology Stack
- **Backend**: Flask (Python 3.11)
- **Authentication**: Flask-Login
- **Security**: Flask-WTF (CSRF Protection)
- **Database**: SQLite3
- **Excel Processing**: openpyxl, pandas
- **Frontend**: HTML, CSS, JavaScript, Bootstrap 5
- **Spreadsheet UI**: Handsontable

## Project Structure
```
.
├── app.py                  # Main Flask application
├── load_excel.py          # Script to load Excel into database
├── checklist.db           # SQLite database (auto-generated)
├── templates/
│   ├── base.html          # Base template with navigation
│   ├── login.html         # Login page
│   ├── admin_dashboard.html  # Admin user management
│   ├── user_dashboard.html   # User dashboard
│   └── checklist.html     # Excel-like checklist interface
├── static/
│   └── css/
│       └── style.css      # Custom styles
└── attached_assets/
    └── CR Check List - Latest Format_1760583521780.xlsx  # Source Excel file
```

## Database Schema
- **users**: id, username, name, email, department, password (hashed), role
- **worksheets**: id, sheet_name (unique), display_order, uploaded_at
- **checklist_data**: id, sheet_name, row_index, col_index, value, updated_at
- **checklist_structure**: id, sheet_name, headers (JSON), total_rows, total_cols, uploaded_at

## Usage
1. Access the application at the provided URL
2. Login with admin or user credentials
3. **Admin**: Manage users through User Management page
4. **All Users**: View and edit checklist through Checklist page
5. Upload Excel files to update checklist data
6. Edit cells inline - changes auto-save to database

## Recent Changes (October 16, 2025)
- Initial project setup with Flask and SQLite
- Implemented user authentication with role-based access
- Created admin panel for user management
- Built Excel file upload and parsing functionality
- Integrated Handsontable for Excel-like editing interface
- Pre-loaded sample checklist data from Excel file
- **Added multi-worksheet support**: Application now displays all worksheets from Excel files as navigable tabs
- **Updated database schema**: Added worksheets table and sheet_name columns to support multiple worksheets
- **Enhanced security**: Added CSRF protection for all authenticated routes using Flask-WTF
- **Added Excel download feature**: Users can now download the current checklist data as an Excel file with the same template format, making it easy to edit offline and maintain the raw Excel file

## Configuration
- Session secret: Uses SESSION_SECRET environment variable (or dev key)
- Server: Runs on 0.0.0.0:5000 in development mode
- Database: SQLite file-based database (checklist.db)

## Security Features
- **Password Hashing**: All passwords are hashed using Werkzeug's secure password hashing
- **CSRF Protection**: Flask-WTF provides CSRF protection for all authenticated routes
- **Role-based Access**: Admin and User roles with restricted access to sensitive operations
- **Session Management**: Secure session handling with Flask-Login
- **Input Validation**: Form validation and sanitization to prevent SQL injection and XSS
