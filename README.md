# Zhanyi Mail

Internal customer email management for `contact@zhanyimetal.com`.

## Features

- Password-protected mailbox dashboard
- Contacts CRUD and CSV import
- Rich-text compose with merge fields
- Durable bulk-send queue with per-recipient status and retry
- Resend delivery and inbound webhooks with Svix signature verification
- Inbox, sent history, replies, and incoming attachment storage
- Inbox conversations grouped by customer with a two-way message timeline
- Automatic inbound forwarding rules with QQ SMTP support
- SiliconFlow / DeepSeek AI drafts using server-selected customer history
- AI draft generation, polishing, and concise rewriting in the compose editor
- Drag-and-drop Excel/CSV lead parsing with import history, duplicate detection, statistics, and searchable frontend results

## Server

- App directory: `/opt/zhanyi-mail`
- Service: `zhanyi-mail.service`
- Local bind: `127.0.0.1:8125`
- Public URL: `https://39.107.111.115`
- Database: `/opt/zhanyi-mail/data/mail.db`
- Configuration: `/opt/zhanyi-mail/.env`

Useful commands:

```bash
systemctl status zhanyi-mail
journalctl -u zhanyi-mail -f
systemctl restart zhanyi-mail
```

## Receiving DNS

Resend receiving requires this record at the root of `zhanyimetal.com`:

```text
Type: MX
Name: @
Priority: 10
Value: inbound-smtp.ap-northeast-1.amazonaws.com
```

The Resend webhook is already created. Receiving becomes active after the MX record verifies.

## AI Configuration

The server `.env` supports:

```text
SILICONFLOW_API_KEY=...
SILICONFLOW_BASE_URL=https://api.siliconflow.cn/v1
SILICONFLOW_MODEL=deepseek-ai/DeepSeek-V3.2
```

The API key is used only by the Flask backend and is never exposed to browser JavaScript.

## Automatic Forwarding SMTP

The forwarding worker uses QQ SMTP when these server-side variables are configured:

```text
FORWARD_SMTP_HOST=smtp.qq.com
FORWARD_SMTP_PORT=465
FORWARD_SMTP_USER=your-account@qq.com
FORWARD_SMTP_AUTH_CODE=your-smtp-authorization-code
```

If the SMTP variables are absent, forwarding falls back to the configured Resend API.
