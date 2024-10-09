# Vinted Sale Tracker

This Python script automates the process of tracking sales from Vinted by fetching sale information (item name and price) from your Gmail inbox and appending the details to a Google Sheets file.

## Features
- Fetches Vinted sale emails from your Gmail inbox using the Gmail API.
- Extracts the sold item name and price from the email content.
- Automatically appends the extracted data to a Google Sheets file.

## Prerequisites

1. **Python 3.x**
2. **Pip** for installing required dependencies.

## Setup

1. **Google Cloud Project Setup:**
   - Enable the **Gmail API** and **Google Sheets API** in your Google Cloud project.
   - Download your OAuth 2.0 credentials (`credentials.json`).

2. **Install Required Libraries:**
   Install the required Python libraries by running:

   ```bash
   pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib beautifulsoup4
```


