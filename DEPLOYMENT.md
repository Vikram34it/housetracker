# House Tracker - Google Sheets Deployment Guide

## Overview

This guide will help you deploy the House Tracker with Google Sheets as the database. This allows multiple users (2-5) to access and update the data simultaneously.

## Prerequisites

- A Google account
- Access to Google Sheets and Google Apps Script

## Step 1: Create a Google Sheet

1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Name it "House Tracker Data" (or any name you prefer)
4. Copy the spreadsheet ID from the URL:
   - Example URL: `https://docs.google.com/spreadsheets/d/ABC123xyz/edit`
   - The ID is: `ABC123xyz`

## Step 2: Set Up Google Apps Script

1. In your Google Sheet, go to **Extensions > Apps Script**
2. Delete any existing code in the editor
3. Copy the contents of `Code.gs` file and paste it into the Apps Script editor
4. Click the **Save** button (floppy disk icon)
5. Name your project "House Tracker Backend"

## Step 3: Deploy the Web App

1. In Apps Script, click **Deploy > New deployment**
2. Click the gear icon and select **Web app**
3. Configure the deployment:
   - **Description**: "House Tracker API"
   - **Execute as**: Me (your email)
   - **Who has access**: Anyone (even unauthenticated)
4. Click **Deploy**
5. Copy the **Web app URL** (it will look like: `https://script.google.com/macros/s/AKfyc.../exec`)

## Step 4: Configure the HTML File

1. Open `house-tracker.html` in a text editor
2. Find this line near the top of the `<script>` section:
   ```javascript
   const APPS_SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE';
   ```
3. Replace `YOUR_APPS_SCRIPT_URL_HERE` with the URL you copied in Step 3
4. Save the file

## Step 5: Initialize the Sheets

1. Open `house-tracker.html` in your browser
2. You will see the login screen
3. Login with default credentials:
   - **Username:** `admin`
   - **Password:** `admin123`
4. The Google Sheet will automatically create four tabs:
   - **Entries**: Stores all points entries
   - **Students**: Stores student academic data
   - **Config**: Stores app configuration
   - **Users**: Stores user accounts and roles

## Step 6: Change Default Admin Password

1. After logging in, note the default password
2. You can change it via the Users management (for now, edit the Apps Script code to change `DEFAULT_ADMIN.password`)

## Step 7: Add Additional Users

1. Click the **Users** tab in the navigation
2. Click **+ Add User**
3. Enter username, full name, password, and role
4. **User roles:**
   - **Admin**: Full access (add/edit/delete entries, manage users)
   - **User**: Can add points and view reports
   - **Viewer**: View only (can view reports but not modify data or add points)

## Step 8: Share the Sheet (Optional)

If you want other users to view the raw data:
1. Open the Google Sheet
2. Click **Share**
3. Add email addresses or set to "Anyone with the link can view"

## Troubleshooting

### "Configure APPS_SCRIPT_URL" error
- Make sure you replaced `YOUR_APPS_SCRIPT_URL_HERE` with the actual Apps Script URL
- The URL should start with `https://script.google.com/macros/s/`

### "Failed to load data" error
- Check that the Apps Script is deployed as a Web app
- Ensure "Who has access" is set to "Anyone"
- Try redeploying the Apps Script

### Data not saving
- Check the browser console for errors (F12 > Console)
- Verify the Apps Script URL is correct
- Make sure you're logged into the Google account that owns the Sheet

### Slow performance
- Google Apps Script has a 6-second execution limit
- For large datasets (1000+ entries), consider optimizing queries

## Security Notes

- The Apps Script URL is public, but it only has access to the specific Google Sheet
- No authentication is required for basic operations
- For sensitive data, consider adding API key authentication

## Updating the Apps Script

If you need to update the backend code:
1. Go to Extensions > Apps Script
2. Make your changes
3. Save the file
4. Deploy > Manage deployments
5. Edit the existing deployment
6. Change **Version** to "New version"
7. Click **Deploy**

## Backup and Restore

The app includes backup/restore functionality:
- **Backup**: Downloads all data as a JSON file
- **Restore**: Uploads a JSON file to restore data

This works with Google Sheets storage and can be used to migrate data between sheets.

## Limitations

- Google Apps Script has a 6-second execution limit
- Maximum 50,000 cells per Google Sheet
- Rate limits apply (concurrent requests may queue)
