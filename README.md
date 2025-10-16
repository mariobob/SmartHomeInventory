# Smart Home Inventory

## Overview
This Google Apps Script system provides two features for Smart Home Inventory:
1. **Autofill**: Automatically populates device information when Model number is entered
2. **Validation**: Checks all items against the database and identifies discrepancies

## Features
### Autofill
- ✅ **Automatic population**: Enter a model number → related fields auto-fill instantly
- ✅ **Smart detection**: Only fills empty fields or fields with "-" placeholder
- ✅ **Non-intrusive**: Won't overwrite data you've already entered
- ✅ **User feedback**: Shows notifications on success or if model not found

### Validation
- ✅ **Comprehensive checks**: Validates all items against the device database
- ✅ **Beautiful HTML report**: Shows discrepancies in an easy-to-read format
- ✅ **Cell references**: Identifies exactly which cells have issues
- ✅ **Auto-fix option**: Automatically correct all discrepancies with one click
- ✅ **Database stats**: View statistics about your device database

### General
- ✅ **Custom menu**: Easy access to all tools via "🏠 Smart Home Tools" menu
- ✅ **Clean code**: Well-structured with separated concerns
- ✅ **70+ device models**: Pre-loaded with existing inventory data

## Installation
### Step 1: Open Google Spreadsheet
Open the Google Spreadsheet containing your "Items" sheet.

### Step 2: Open Apps Script editor
1. In Google Sheets, click **Extensions** → **Apps Script**
2. This opens the script editor in a new tab

### Step 3: Create script files
Create **5 files** in the Apps Script editor and copy their contents and names.
Delete the default `Code.gs` file if it still exists.

### Step 4: Name the project
1. Click "Untitled project" at the top
2. Rename it to "Smart Home Inventory - Autofill"
3. Click away to save the name

### Step 5: Reload spreadsheet
1. Close the Apps Script tab
2. Go back to the spreadsheet
3. Reload the page
4. A new menu should appear: **🏠 Smart Home Tools**

### Step 6: Authorize the Script
1. Click **🏠 Smart Home Tools** → **✅ Validation** → **Show Database Stats**
2. Google will ask for authorization
3. Click **Review Permissions** → Choose your account
4. Click **Advanced** → **Go to Smart Home Inventory - Autofill**
5. Click **Allow**
