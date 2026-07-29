# Viva House Championship Tracker

A live inter-house points tracker that uses Google Sheets as a database. Supports multiple users (2-5) with real-time synchronization.

## Features

- Real-time data sync via Google Sheets
- **User authentication** with admin and user roles
- Excel/CSV import for student marks
- Dashboard with podium, charts, and analytics
- Monthly statistics view
- Awards and records
- Activity history with search and filters
- Backup/restore functionality
- Mobile responsive design

## Login & User Roles

### Default Admin Account
- **Username:** `admin`
- **Password:** `admin123`

**Important:** Change the default password after first login!

### User Roles

| Role | Permissions |
|------|-------------|
| **Admin** | Full access: Add/edit/delete entries, manage users, import Excel, reset data |
| **User** | View only: Can view dashboard, reports, and history (no editing) |

### Managing Users (Admin Only)

1. Login as admin
2. Click the **Users** tab in the navigation
3. Add new users with username, full name, password, and role
4. Edit or delete existing users (except the main admin)

## Live Demo

[Deploy your own](#deployment) or check the live version (once deployed)

## Quick Start

### Prerequisites

- A Google account
- Git installed on your computer

### Deployment Steps

#### 1. Set Up Google Sheets Backend

1. Create a new [Google Sheet](https://sheets.google.com)
2. Go to **Extensions > Apps Script**
3. Copy the contents of `Code.gs` into the script editor
4. Click **Deploy > New deployment**
5. Select **Web app** and configure:
   - Execute as: Me
   - Who has access: Anyone
6. Copy the deployment URL

#### 2. Configure the HTML

1. Open `house-tracker.html`
2. Find this line:
   ```javascript
   const APPS_SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE';
   ```
3. Replace with your Apps Script URL

#### 3. Deploy to GitHub Pages

1. Create a new repository on GitHub
2. Upload these files:
   - `house-tracker.html`
   - `VIVAA.png`
   - `xlsx.full.min.js`
   - `house-marks-template.csv`

3. Go to **Settings > Pages**
4. Select **Source: Deploy from a branch**
5. Choose **main** branch and **/ (root)** folder
6. Click **Save**

Your app will be live at: `https://yourusername.github.io/your-repo-name/`

## Files Overview

| File | Description |
|------|-------------|
| `house-tracker.html` | Main application |
| `Code.gs` | Google Apps Script backend |
| `VIVAA.png` | Logo image |
| `xlsx.full.min.js` | Excel parsing library |
| `house-marks-template.csv` | Import template |
| `test-backend.html` | Backend testing tool |
| `DEPLOYMENT.md` | Detailed deployment guide |

## Configuration

Edit the `APPS_SCRIPT_URL` in `house-tracker.html`:

```javascript
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec';
```

## Usage

### For Admins
1. **Adding Points**: Click "+ Add Points" button
2. **Bulk Add**: Click "Bulk" button for multiple houses
3. **Excel Import**: Use the "Excel Import" tab to import student marks
4. **Manage Users**: Click "Users" tab to add/edit/delete users
5. **View Stats**: Navigate through Dashboard, Houses, Comparison, Analytics, Monthly, Awards, and History tabs

### For Users
1. **View Dashboard**: See current standings and podium
2. **View Reports**: Check Houses, Comparison, Analytics, Monthly, Awards tabs
3. **View History**: See activity log (read-only)
4. **Export Data**: Download CSV or backup files

## Security Notes

- The default admin password should be changed immediately
- User passwords are stored as hashed values in Google Sheets
- Users can only view data, not modify it
- Admins have full control over all features
- The Apps Script URL is public but only accesses your specific spreadsheet

## Browser Support

- Chrome (recommended)
- Firefox
- Safari
- Edge

## License

MIT License - feel free to use and modify for your school.
