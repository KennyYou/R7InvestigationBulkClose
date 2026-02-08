# InsightIDR Investigation Updater


Rapid7 does not natively provide a way to apply individual comments while bulk-closing investigations. The built-in bulk close feature closes on investigations that share the same name, which is not suitable for scenarios where alerts are distinct, unrelated issues.

This app provides an local GUI that interfaces with the Rapid7 API to view, update, and more importantly update multiple investigations without needing to manually close them out one by one.

Notes
If an ORG ID is not configured, direct links to individual investigations will be unavailable. After setting the ORG ID, a restart of the application may be required for the changes to take effect.

<img width="2394" height="1264" alt="Screenshot 2026-02-06 191629" src="https://github.com/user-attachments/assets/f8cd2193-e2ce-4972-8b21-36e10556bc7c" />

You can check if an investigation has an comment by selecting the checkbox and selecting 'refresh' on the comments sections. 

## Requirements

- Rapid7 InsightIDR API key (Each member will need permissions to generate their own, otherwise the comments and audit trail will point to however initally generated the event)
- Know your Org ID & Region R7 is hosted in (us / us3 / etc)

## Notes

- On first run, the app prompts for:
  - Region + Organization settings
  - Team member assignees
  - API key source (environment variable or key file)
  
  <img width="835" height="505" alt="image" src="https://github.com/user-attachments/assets/c5a6a003-c472-4086-8208-ba2afdafcc6a" />

