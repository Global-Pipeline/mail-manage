# Data Snapshot

This directory contains a private production SQLite snapshot created with SQLite's online backup command.

- Snapshot time: `2026-08-23 16:20:24 +08:00`
- Database: `mail.db`
- Attachments: `attachments/`

The database contains customer correspondence and contact information. Keep the repository private and restrict access accordingly.

To restore it on the server, stop the application, back up the current database, replace `/opt/zhanyi-mail/data/mail.db`, preserve ownership for the `zhanyimail` user, and restart the service.
