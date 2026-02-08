<div align="center">

# InsightIDR Investigation Updater

**Bulk-update Rapid7 InsightIDR investigations**
</div>

For Background.

Rapid7 does not natively provide a way to apply individual comments while bulk-closing investigations. Their built-in bulk close feature closes on investigations that share the same name, which is not ideal for scenarios where alerts are essentially distinct, unrelated issues. Nor does it add a comment to why this was closed if you decide to bulk close one. This app provides an local GUI that interfaces with the Rapid7 API to view, update, and more importantly update multiple investigations without needing to manually close them out one by one.



<img width="2394" height="1264" alt="Screenshot 2026-02-06 191629" src="https://github.com/user-attachments/assets/872b53ef-8929-48be-8591-8f5c9b6ff945" />


---

## Table of Contents

- [Key Features](#key-features)
- [Usage](#usage)
- [Troubleshooting](#troubleshooting)

---

## Key Features

- **Bulk updates** for selected investigations:
  - status
  - disposition
  - assignee
  - optional comment
- **Assignee management UI** (add/edit/remove team members for assignment).
- **Region and organization settings** with first-run setup dialog.
- **API key source options**:
  - environment variable
  - local key file
- **Comments workflow**:
  - post comments to investigations
  - view selected investigation comments
  - local reusable comment history

On first launch, the app guides you through:

1. Region + Organization settings
2. Team member (assignee) setup
3. API key source selection

---

## Usage

1. Click **Refresh** to load open-like investigations.
2. Select one or more investigations.
3. Choose desired **Status**, **Disposition**, and/or **Assignee**.
4. Add an optional comment.
5. Click **Update Selected**.

### Right panel tabs

- **Status Log** – operation and error messages
- **Comments** – view comments for selected investigation (You will need to select one ivnestigation then hit refresh)
- **History** – reusable saved comments

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


### Something else just open an issue

- If something else is broken just open an issue and I can take a look.


