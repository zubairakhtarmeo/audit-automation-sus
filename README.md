# Audit Automation Tool

AI-powered desktop-style Streamlit app for auto-filling brand compliance/audit Excel files (Bestseller, Jack & Jones, Kiabi, etc.) using Anthropic Claude.

## Run locally

1. Create `.env` with your key:

```
ANTHROPIC_API_KEY=YOUR_KEY
```

2. Install dependencies:

```
pip install -r requirements.txt
```

3. Start the app:

```
streamlit run ui/app.py
```

## Notes
- The API key is read from environment / `.env` and is not stored permanently.
- Output files are saved as `Filled_[original].xlsx`.
