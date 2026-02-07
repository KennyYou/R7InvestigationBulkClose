# InsightIDR Investigation Updater

Desktop bulk updater for Rapid7 InsightIDR investigations using `customtkinter`.

## Requirements

- Python 3.10+
- Rapid7 InsightIDR API key
- Network access to `*.api.insight.rapid7.com`

Install Python dependencies:

```bash
pip install -r requirements.txt
```

## Run

```bash
python insightidr_updater.py
```

## Notes

- On first run, the app prompts for:
  - Region + Organization settings
  - Team member assignees
  - API key source (environment variable or key file)
- App settings are persisted under the user config directory in `InsightIDRUpdater/config.json`.
