# Google Drive API Credentials Setup

This directory contains credentials for connecting to Google Drive API.

## Setup Instructions

### Step 1: Go to Google Cloud Console
Navigate to [Google Cloud Console](https://console.cloud.google.com/) and create a new project or select an existing one.

### Step 2: Enable Google Drive API
1. Go to **APIs & Services** > **Library**
2. Search for "Google Drive API"
3. Click on **Google Drive API** and press **Enable**

### Step 3: Create OAuth 2.0 Credentials
1. Go to **APIs & Services** > **Credentials**
2. Click **Create Credentials** > **OAuth client ID**
3. Select **Desktop app** as the application type
4. Give it a name (e.g., "Randomizador de Pruebas")
5. Click **Create**

### Step 4: Download Credentials
1. Download the JSON file from the OAuth client you just created
2. Save it as `credentials/client_secret.json` in the project root

### Step 5: First Run
On first run, the application will:
- Open a browser window for authentication
- Generate `credentials/token.json` automatically after you grant permission

## Important Notes

- Never commit `client_secret.json` or `token.json` to version control
- `token.json` is automatically generated and refreshed as needed
- If you need to re-authenticate, delete `token.json` and run the application again

## Security Warning

The `.gitignore` file is configured to exclude:
- `credentials/client_secret.json`
- `credentials/token.json`

Do not remove these entries from `.gitignore`.