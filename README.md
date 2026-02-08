<div align="center">

# InsightIDR Investigation Updater

**Bulk-update Rapid7 InsightIDR investigations from a desktop UI built with CustomTkinter.**

![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-8A2BE2)
![UI](https://img.shields.io/badge/UI-CustomTkinter-2ea44f)

</div>

---

## Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [Requirements](#requirements)
- [Quick Start](#quick-start)
- [Configuration](#configuration)
- [Usage](#usage)
- [Data & Privacy Notes](#data--privacy-notes)
- [Troubleshooting](#troubleshooting)

---

## Overview

This tool helps SOC and IR teams quickly triage and update InsightIDR investigations in bulk.
It supports status/disposition/assignee updates, comment posting, filtering, and a right-side activity panel for status, comments, and comment history.

---

## Key Features

- **Bulk updates** for selected investigations:
  - status
  - disposition
  - assignee
  - optional comment
- **Assignee management UI** (add/edit/remove team members).
- **Region and organization settings** with first-run setup dialog.
- **API key source options**:
  - environment variable
  - local key file
- **Comments workflow**:
  - post comments to investigations
  - view selected investigation comments
  - local reusable comment history
- **Responsive UX** with background threads and progress dialogs.

---

## Requirements

- Python **3.10+**
- Rapid7 InsightIDR API key
- Network access to `*.api.insight.rapid7.com`

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Quick Start

Run the app:

```bash
python insightidr_updater.py
```

On first launch, the app guides you through:

1. Region + Organization settings
2. Team member (assignee) setup
3. API key source selection

---

## Configuration

The app stores settings in your user config directory:

- **Windows:** `%APPDATA%/InsightIDRUpdater/config.json`
- **macOS:** `~/Library/Application Support/InsightIDRUpdater/config.json`
- **Linux:** `~/.config/InsightIDRUpdater/config.json`

Stored settings include items such as:

- API key source metadata (env var name or key file path)
- region and optional org ID
- assignees
- comment history

---

## Usage

1. Click **Refresh** to load open-like investigations.
2. Select one or more investigations.
3. Choose desired **Status**, **Disposition**, and/or **Assignee**.
4. Add an optional comment.
5. Click **Update Selected**.

### Right panel tabs

- **Status Log** – operation and error messages
- **Comments** – view comments for selected investigation
- **History** – reusable saved comments

---

## Data & Privacy Notes

- This tool can display and store operational data (e.g., assignee emails, org ID, and comment history) in local settings.
- Protect workstation access and local config files appropriately.
- Avoid sharing screenshots that may reveal sensitive investigation details.

---

## Troubleshooting

### API key not detected

- Confirm the configured env var exists in your shell/session, or
- Confirm your key file path points to a file containing only the API key.

### No investigations shown

- Verify region setting matches your InsightIDR region.
- Confirm API key has correct permissions.
- Check connectivity to `*.api.insight.rapid7.com`.

### Comments fail to load/post

- Ensure investigation RRN can be resolved.
- Verify API key scope includes v1 comment access.

---

If helpful, I can also add screenshots/GIFs and a concise architecture section to make this README even closer to high-end showcase-style repos.
