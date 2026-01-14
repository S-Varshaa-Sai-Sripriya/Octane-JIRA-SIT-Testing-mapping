# Octane ID - JIRA ID Mapper

A desktop application that transforms Excel files by mapping Octane IDs to JIRA IDs, creating separate rows for each JIRA ID mapping.

## Overview

This tool processes Excel files containing test execution data and creates a mapped output where each JIRA ID gets its own row. 

## Features

- 📂 **Easy File Upload**: Browse and select your input Excel file
- ⚡ **Fast Processing**: Processes entire Excel files in seconds
- 👁️ **Preview**: See the first 10 rows of output before downloading
- 💾 **Excel Export**: Download transformed data as Excel (.xlsx) format (Use full screen to see the download button)
- 🖥️ **User-Friendly GUI**: Simple desktop interface built with tkinter

## Prerequisites

- Python 3.7 or higher
- pip (Python package installer)

## Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/S-Varshaa-Sai-Sripriya/Octane-JIRA-SIT-Testing-mapping.git
   cd Octane-JIRA-SIT-Testing-mapping
   ```

2. **Install required dependencies**
   ```bash
   pip install openpyxl pandas
   ```

## Usage

### Running the Application

Run the desktop GUI application:

```bash
python gui_app.py
```

### Step-by-Step Instructions

1. **Launch the application**
   - Run `python gui_app.py` in your terminal
   - A window titled "Octane ID - JIRA ID Mapper" will open

2. **Select Input File**
   - Click the **"Browse"** button
   - Select your Excel file (must contain columns: "Test Team", "ID", "Test: JIRA ID")

3. **Process the Data**
   - Click **"Compute Mapping"** button
   - Wait for processing to complete (progress bar will show)

4. **Preview Results**
   - View the first 10 rows in the preview table
   - Check that the mapping is correct

5. **Download Output**
   - Click **"📥 Download Output Excel"** button (appears after processing in full screen)
   - Choose location and filename for the output file
   - Save the transformed Excel file

## Requirements

```
openpyxl>=3.0.0
pandas>=1.3.0
```


