# Invoice Assistant (using Agno)

**Invoice Assistant** is a lightweight AI-powered tool that helps you manage timesheets and generate invoices with ease. It leverages **Agno AI** — a personalized AI that runs locally, and learns from your interactions. This assistant supports Telegram integration for easy messaging and notifications.


## 🧠 What is Agno AI?

[Agno AI](https://www.agno.ai/) is a privacy-first personal AI framework designed to run on your device. It remembers your interactions, works offline, and can integrate with external tools like Telegram. This project demonstrates one of Agno’s use cases: creating a context-aware invoice assistant that learns from your past input stored in a local memory database.


## 🚀 Features

* Parses timesheet data from Excel files
* Generates invoices automatically
* Maintains user memory using SQLite
* Supports Telegram notifications via a bot
* Powered by Agno AI, built for local-first personal intelligence


## 📁 Project Structure

```
.
├── pyproject.toml          # Project metadata and build system config
├── requirements.txt        # Project dependencies
├── timesheet_july.xlsx     # Sample timesheet input
├── tools.py                # Utility functions
├── invoice_assist.py       # Main application entry point
```


## 🔧 Setup Instructions

### 1. Clone the Repository

### 2. Install Dependencies with `uv`

Install dependencies using [uv](https://github.com/astral-sh/uv), a fast Python package manager:

```bash
uv sync
uv pip install requirements.txt
```

### 3. Create a `.env` File

Create a `.env` file in the root directory with the following keys:

```
GOOGLE_API_KEY=your_google_api_key
TELEGRAM_BOT_TOKEN=your_bot_token
TELEGRAM_CHAT_ID=your_telegram_chat_id
```

* ### **To get your Telegram Chat ID**: Message `@userinfotelegram` on Telegram and send `/start` to retrieve it.
* ### **Need a bot token?** DM me for my bot token.


### 4. Run the Assistant

```bash
uv run invoice_assist.py
```

On first run, the app will create a `memory/` directory with a SQLite file to store user memory and contextual data for Agno AI.


## ➕ Adding New Packages

To add a new dependency:

```bash
uv pip install <package-name>
```

Then update the requirements file:

```bash
uv pip freeze > requirements.txt
```

## 📬 Support / Feedback

If you have questions, suggestions, or need help setting up your Telegram bot, feel free to reach out.

