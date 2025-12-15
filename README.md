# G-Drive Audit Logger

A robust Google Apps Script solution for auditing Google Drive activity in corporate environments. 

## The Problem
Google Drive Activity API provides granular data about file changes but often returns opaque user identifiers (e.g., `people/123456789`) instead of readable names. Additionally, raw logs can be overwhelming and difficult to analyze over time.

## The Solution
This script acts as an automated audit logger that:
1. **Fetches Activity:** Queries Drive Activity API v2 for file changes (Edit, Create, Move, Rename).
2. **Normalizes Users:** Maps `people/ID` to human-readable names using a configurable `users` sheet.
3. **Smart Caching:** Caches folder structures to minimize Drive API calls and avoid quota limits.
4. **Incremental Logging:** Stores logs in daily sheets (`log_YYYY-MM-DD`) to prevent "black hole" datasets.
5. **Self-Healing:** Includes utilities to retroactively fix user names in old logs if the user directory is updated.

## Setup

1. Create a new Google Apps Script project.
2. Enable **Drive Activity API** and **Google Drive API** in the *Services* tab.
3. Copy the code files into the project.
4. Update `ROOT_FOLDER_ID` in `Code.gs` with the ID of the folder you want to monitor.
5. Run `logDriveActivity()` manually once to initialize the Spreadsheet.
6. Set up a time-based trigger (e.g., every hour) to run `logDriveActivity()`.

## User Mapping
The script automatically creates a `users` sheet in the log spreadsheet. Populate it to resolve IDs:
| actorPersonName | displayName |
|-----------------|-------------|
| people/12345... | Alice Smith |
| people/67890... | Bob Jones   |

## License
MIT
