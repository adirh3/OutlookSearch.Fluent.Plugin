# Outlook Search — Fluent Search Plugin

A plugin for [Fluent Search](https://fluentsearch.net/) that lets you search and interact with Microsoft Outlook emails and calendar events directly from the launcher.

## Features

### Email Search
- Search emails by subject, sender, body, or recipients
- Results display subject, sender, relative time, attachment indicator, and body preview snippet
- Rich preview panel showing From, To, CC, Date, Folder, Importance, and body preview
- Unread/read status shown via distinct icons

### Calendar Search
- Search calendar events by title, location, or body content
- Shows event time range, location, organizer, and attendees
- Smart status indicators ("Happening now", "In 30 minutes", etc.)
- Empty search with `outlook` tag shows upcoming events

### Quick Actions
- **Compose New Email** — Opens a new email compose window
- **Schedule New Meeting** — Opens meeting scheduling in Outlook
- **Open Inbox** — Opens the Outlook inbox
- **Open Calendar** — Opens the Outlook calendar
- **Sign Out** — Sign out and clear cached credentials

### Operations
- **Open Email / Open Event** — Open in your browser (Outlook Web)
- **Reply** / **Reply All** / **Forward** — Direct email actions (COM backend)
- **Copy Subject / Copy Title** — Copy to clipboard

### Search Tags
| Tag | Description |
|---|---|
| `outlook` | Search both emails and events, plus quick actions |
| `email` | Search emails only |
| `calendar` | Search calendar events only |

Email results show the `email` + `outlook` tags. Event results show the `calendar` + `outlook` tags.

## Backends

The plugin supports two backends and uses whichever is available:

### Microsoft Graph API (New Outlook / Outlook.com)
- Sign in with your Microsoft account via the plugin
- Works with personal (outlook.com) and work/school (Microsoft 365) accounts
- Token is cached locally so you stay signed in across restarts
- Requires internet connectivity

### COM Interop (Classic Outlook desktop app)
- Automatic — no sign-in needed if classic Outlook is installed and running
- 100% local, no data leaves your machine, works offline
- Supports Reply, Reply All, Forward, and opening items directly in Outlook

If both backends are available, COM interop is preferred for email operations (direct Outlook actions), while Graph is used for search.

## Settings

| Setting | Default | Description |
|---|---|---|
| Maximum email results | 15 | Max email results per search |
| Maximum event results | 10 | Max event results per search |
| Search emails from last N days | 90 | How far back to search emails |
| Search events from last N days | 30 | How far back to search past events |
| Search events N days ahead | 90 | How far ahead to search future events |
| Show quick actions | Yes | Show quick action results |
| Search emails | Yes | Include email results |
| Search calendar events | Yes | Include calendar event results |
| Show upcoming events on empty | Yes | Show upcoming events with empty search |
| Microsoft Graph Client ID | _(built-in)_ | Custom Azure AD app ID (advanced) |

## Requirements

- Windows
- **New Outlook / Outlook.com**: internet connection, one-time sign-in
- **Classic Outlook**: Office 2016+ desktop app installed

## Build

```
dotnet build OutlookSearch.Fluent.Plugin.sln -c Release
```

## Test

```
dotnet test OutlookSearch.Fluent.Plugin.sln -c Release
```
