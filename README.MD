# Power BI Documenter

A Streamlit app that automatically generates documentation for Power BI reports using the Power BI Admin API and GPT-4o. Exports structured documentation to Excel or Word.

## Overview

This tool connects to your Power BI tenant or accepts a `.pbit` file directly, extracts report metadata (tables, columns, measures, and data sources), and uses GPT-4o to produce human-readable documentation — exported as a formatted `.xlsx` or `.docx` file.

## Requirements

- Python 3.8+
- An [OpenAI](https://platform.openai.com/) API key with GPT-4o access
- A Microsoft Azure App Registration with Power BI Admin API permissions (for workspace mode)

## Setup

### 1. Clone the repository

```bash
git clone <repo-url>
cd <repo-folder>
```

### 2. Install dependencies

```bash
pip install streamlit msal requests pandas openai python-docx xlsxwriter gspread google-auth python-dotenv
```

### 3. Configure environment variables

Create a `.env` file in the project root:

```env
API_KEY=your_openai_api_key
```

### 4. Run the app

```bash
streamlit run app.py
```

## Usage

### Option A — Connect via Azure App

Fill in the sidebar with your Azure App credentials:

- **App ID** — Application (client) ID from Azure Portal
- **Tenant ID** — Your Microsoft tenant ID
- **Secret value** — Client secret from the App Registration

The app will list all available workspaces and reports in your Power BI tenant.

### Option B — Upload a `.pbit` file

Upload a Power BI Template file directly via the sidebar. No Azure credentials needed.

> To export a `.pbit` file: in Power BI Desktop go to **File → Export → Power BI Template**.

### Generating documentation

Once a report is loaded, the following actions are available:

| Action | Output |
|---|---|
| Download full table as Excel | Raw extracted metadata as `.xlsx` |
| Document report to Excel | AI-generated docs across 4 sheets: report info, tables, measures, data sources |
| Document report to Word | AI-generated docs as a `.docx` file |

You can also preview filtered views inline — tables only, columns only, or measures only — before downloading.

## Azure App Setup

To use workspace mode, your Azure App Registration needs the following Power BI API permission:

- `Tenant.Read.All` or `Tenant.ReadWrite.All`

The app authenticates using the **client credentials flow** — no user login is required.

## Rate Limits

The Power BI Admin API is limited to **50 requests per hour** or **15 per minute**. The app handles HTTP 429 responses automatically using exponential backoff, but you may need to wait before making new requests if the limit is reached.

## Project Structure

```
.
├── app.py    # Main Streamlit application
└── .env      # Environment variables (not committed)
```

## Environment Variables

| Variable  | Description             | Required |
|-----------|-------------------------|----------|
| `API_KEY` | OpenAI API key (GPT-4o) | Yes      |
