# M365 CLI

Modern command-line interface for Microsoft 365 personal accounts (Mail, Calendar, OneDrive, SharePoint).

## Features

- 📧 **Mail**: List, read, send, search emails (with attachments)
- 📅 **Calendar**: Manage events (list, create, update, delete)
- 📁 **OneDrive**: File management (upload, download, search, share)
- 🌐 **SharePoint**: Site and document management
- 🔐 **Secure**: OAuth 2.0 Device Code Flow authentication
- 🚀 **Fast**: Minimal dependencies, uses native Node.js fetch
- 🤖 **AI-friendly**: Clean text output + JSON option

## Installation

```bash
cd ~/Projects/m365-cli
npm install
npm link
```

After linking, the `m365` command will be available globally.

## Quick Start

### 1. Login

```bash
m365 login
```

Follow the prompts to authenticate with your Microsoft 365 account using Device Code Flow.

### 2. List Emails

```bash
m365 mail list --top 5
```

### 3. View Calendar

```bash
m365 calendar list --days 7
```

### 4. Browse OneDrive

```bash
m365 onedrive ls
```

## Commands Reference

### Authentication

```bash
m365 login              # Login with Device Code Flow
m365 logout             # Clear stored credentials
```

### Mail Commands

```bash
# List emails
m365 mail list [options]
  --top <n>                         # Number of emails (default: 10)
  --folder <name>                   # Folder to list (default: inbox)
  --json                            # Output as JSON

# Supported folders:
#   inbox       - Inbox (default)
#   sent        - Sent Items
#   drafts      - Drafts
#   deleted     - Deleted Items
#   junk        - Junk Email
#   Or use a folder ID directly

# Read email
m365 mail read <id> [options]
  --json                            # Output as JSON

# Send email
m365 mail send <to> <subject> <body> [options]
  --attach <file1> <file2> ...      # Attach files
  --json                            # Output as JSON

# Search emails
m365 mail search <query> [options]
  --top <n>                         # Number of results (default: 10)
  --json                            # Output as JSON

# List attachments
m365 mail attachments <id> [options]
  --json                            # Output as JSON

# Download attachment
m365 mail download-attachment <message-id> <attachment-id> [local-path] [options]
  --json                            # Output as JSON
```

**Examples:**
```bash
m365 mail list --top 5
m365 mail list --folder sent --top 10    # List sent emails
m365 mail list --folder drafts           # List draft emails
m365 mail read AAMkADA5ZDE2Njk2...
m365 mail send "user@example.com" "Meeting" "Let's meet tomorrow" --attach report.pdf
m365 mail search "project update" --top 20

# Attachment examples
m365 mail attachments AAMkADA5ZDE2Njk2...      # List all attachments in an email
m365 mail download-attachment AAMkADA5... AAMkAGQ...   # Download using attachment's original filename
m365 mail download-attachment AAMkADA5... AAMkAGQ... ~/Downloads/file.pdf  # Specify output path
```

### Calendar Commands

```bash
# List events
m365 calendar list [options]
m365 cal list [options]             # Alias
  --days <n>                        # Look ahead N days (default: 7)
  --top <n>                         # Maximum events (default: 50)
  --json                            # Output as JSON

# Get event details
m365 calendar get <id> [options]
  --json                            # Output as JSON

# Create event
m365 calendar create <title> [options]
  --start <datetime>                # Start time (required)
  --end <datetime>                  # End time (required)
  --location <location>             # Event location
  --body <description>              # Event description
  --attendees <email1,email2>       # Attendee emails (comma-separated)
  --allday                          # All-day event
  --json                            # Output as JSON

# Update event
m365 calendar update <id> [options]
  --title <title>                   # New title
  --start <datetime>                # New start time
  --end <datetime>                  # New end time
  --location <location>             # New location
  --body <description>              # New description
  --json                            # Output as JSON

# Delete event
m365 calendar delete <id> [options]
  --json                            # Output as JSON
```

**Datetime formats:**
- Full datetime: `2026-02-17T14:00:00`
- Date only (for all-day events): `2026-02-17`

**Examples:**
```bash
m365 cal list --days 5
m365 cal get AAMkADA5ZDE2Njk2...
m365 cal create "Team Meeting" --start "2026-02-17T14:00:00" --end "2026-02-17T15:00:00" --location "Room A"
m365 cal create "Holiday" --start "2026-02-20" --end "2026-02-21" --allday
m365 cal update AAMkADA5... --location "Room B"
m365 cal delete AAMkADA5...
```

### OneDrive Commands

```bash
# List files
m365 onedrive ls [path] [options]
m365 od ls [path] [options]         # Alias
  --top <n>                         # Maximum items (default: 100)
  --json                            # Output as JSON

# Get file/folder metadata
m365 onedrive get <path> [options]
  --json                            # Output as JSON

# Download file
m365 onedrive download <remote-path> [local-path] [options]
  --json                            # Output as JSON

# Upload file
m365 onedrive upload <local-path> [remote-path] [options]
  --json                            # Output as JSON

# Search files
m365 onedrive search <query> [options]
  --top <n>                         # Maximum results (default: 50)
  --json                            # Output as JSON

# Create sharing link
m365 onedrive share <path> [options]
  --type <view|edit>                # Link type (default: view)
  --json                            # Output as JSON

# Create folder
m365 onedrive mkdir <path> [options]
  --json                            # Output as JSON

# Delete file/folder
m365 onedrive rm <path> [options]
  --force                           # Skip confirmation
  --json                            # Output as JSON
```

**Examples:**
```bash
m365 od ls
m365 od ls Documents
m365 od get "Documents/report.pdf"
m365 od download "Documents/report.pdf" ~/Downloads/
m365 od upload ~/Desktop/photo.jpg "Photos/vacation.jpg"
m365 od search "budget" --top 20
m365 od share "Documents/report.pdf" --type edit
m365 od mkdir "Projects/New Project"
m365 od rm "old-file.txt" --force
```

**Features:**
- 🚀 Large file upload support (automatic chunking for files ≥4MB)
- 📊 Progress display for downloads and uploads
- 💾 Human-readable file sizes (B/KB/MB/GB/TB)
- ⚠️ Confirmation prompt before deletion (unless `--force`)

### SharePoint Commands

```bash
# List sites
m365 sharepoint sites [options]
m365 sp sites [options]             # Alias
  --search <query>                  # Search for sites
  --top <n>                         # Maximum sites (default: 50)
  --json                            # Output as JSON

# List site lists
m365 sharepoint lists <site> [options]
  --top <n>                         # Maximum lists (default: 100)
  --json                            # Output as JSON

# List list items
m365 sharepoint items <site> <list> [options]
  --top <n>                         # Maximum items (default: 100)
  --json                            # Output as JSON

# List files in document library
m365 sharepoint files <site> [path] [options]
  --top <n>                         # Maximum files (default: 100)
  --json                            # Output as JSON

# Download file
m365 sharepoint download <site> <file-path> [local-path] [options]
  --json                            # Output as JSON

# Upload file
m365 sharepoint upload <site> <local-path> [remote-path] [options]
  --json                            # Output as JSON

# Search content
m365 sharepoint search <query> [options]
  --top <n>                         # Maximum results (default: 50)
  --json                            # Output as JSON
```

**Site identifier formats:**

SharePoint commands accept sites in multiple formats:

1. **Path format** (recommended):
   ```
   hostname:/sites/sitename
   # Example: contoso.sharepoint.com:/sites/team
   ```

2. **Site ID format** (from `m365 sp sites` output):
   ```
   hostname,siteId,webId
   # Example: contoso.sharepoint.com,8bfb5166-...,ea772c4f-...
   ```

3. **URL format** (for some commands):
   ```
   https://hostname/sites/sitename
   # Example: https://contoso.sharepoint.com/sites/team
   ```

**Tip:** Run `m365 sp sites --json` to get the exact site ID for any site.

**Examples:**
```bash
m365 sp sites
m365 sp sites --search "marketing"

# Using path format (recommended)
m365 sp lists "contoso.sharepoint.com:/sites/team"

# Using site ID from sites list output
m365 sp lists "contoso.sharepoint.com,8bfb5166-7dff-4f15-8d44-3417d68f519a,ea772c4f-..."

m365 sp files "contoso.sharepoint.com:/sites/team" "Documents"
m365 sp download "contoso.sharepoint.com:/sites/team" "Documents/file.pdf"
m365 sp upload "contoso.sharepoint.com:/sites/team" ~/report.pdf "Documents/report.pdf"
m365 sp search "quarterly report" --top 20
```

**Note:** SharePoint commands require `Sites.ReadWrite.All` permission. If you encounter permission errors, run `m365 logout` then `m365 login` to re-authenticate with updated permissions.

## Configuration

### Credentials Location

Credentials are stored at: `~/.openclaw/workspace/creds/.m365-creds`

**Security:** File permissions are set to `600` (owner read/write only). Do not share this file.

### Azure AD Application

This tool uses the following Azure AD configuration:
- **Tenant ID**: `5b4c4b46-4279-4f19-9e5d-84ea285f9b9c`
- **Client ID**: `091b3d7b-e217-4410-868c-01c3ee6189b6`

To use your own Azure AD app:
```bash
export M365_TENANT_ID="your-tenant-id"
export M365_CLIENT_ID="your-client-id"
```

### Permissions

The application requests the following Microsoft Graph permissions:
- `Mail.ReadWrite` - Read and write mail
- `Mail.Send` - Send mail
- `Calendars.ReadWrite` - Read and write calendar events
- `Files.ReadWrite.All` - Read and write files in OneDrive
- `Sites.ReadWrite.All` - Read and write SharePoint sites

## Output Formats

### Default (Text)

Clean, human-readable output with emoji icons:

```
📧 Mail List (top 3)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[1] 📩 Meeting Reminder
    From: alice@example.com
    Date: 2026-02-16 09:30
    ID: AAMkAG...
```

### JSON (--json)

Structured output for scripting and AI consumption:

```json
[
  {
    "id": "AAMkAG...",
    "subject": "Meeting Reminder",
    "from": {
      "emailAddress": {
        "name": "Alice",
        "address": "alice@example.com"
      }
    },
    "receivedDateTime": "2026-02-16T09:30:00Z",
    "isRead": false
  }
]
```

## Project Structure

```
m365-cli/
├── bin/
│   └── m365.js              # CLI entry point
├── src/
│   ├── auth/                # Authentication
│   │   ├── token-manager.js # Token storage & refresh
│   │   └── device-flow.js   # Device Code Flow
│   ├── graph/               # Graph API client
│   │   └── client.js        # HTTP client with auto-refresh
│   ├── commands/            # Command implementations
│   │   ├── mail.js          # Mail commands
│   │   ├── calendar.js      # Calendar commands
│   │   ├── onedrive.js      # OneDrive commands
│   │   └── sharepoint.js    # SharePoint commands
│   └── utils/               # Utilities
│       ├── config.js        # Config management
│       ├── output.js        # Output formatting
│       └── error.js         # Error handling
├── config/
│   └── default.json         # Default configuration
├── package.json             # Project metadata
├── README.md                # This file
└── PLAN.md                  # Development roadmap
```

## Development

### Tech Stack

- **Node.js 18+** - ESM modules, native fetch API
- **commander.js** - CLI framework
- **Microsoft Graph API** - M365 services

### Local Testing

```bash
# Link for local development
npm link

# Test commands
m365 --help
m365 mail --help
m365 mail list --top 3
```

## Troubleshooting

### "Not authenticated" error

Run `m365 login` to authenticate.

### Token expired

Tokens refresh automatically. If refresh fails, run `m365 login` again.

### Permission denied (SharePoint)

SharePoint requires additional permissions. Run:
```bash
m365 logout
m365 login
```

### Network errors

- Check internet connection
- Verify firewall settings
- Ensure Microsoft Graph API is accessible

## Roadmap

- [x] **Phase 1**: Framework + Mail ✅
- [x] **Phase 2**: Calendar ✅
- [x] **Phase 3**: OneDrive ✅
- [x] **Phase 3.5**: SharePoint ✅
- [ ] **Phase 4**: Contacts & Advanced Features
- [ ] **Phase 5**: Optimization & Release

See [PLAN.md](./PLAN.md) for detailed roadmap and implementation notes.

## Security

- ✅ Credentials stored with `600` permissions
- ✅ OAuth 2.0 Device Code Flow
- ✅ Automatic token refresh
- ✅ No sensitive data in logs
- ✅ HTTPS-only communication

## License

MIT License - see LICENSE file for details.

## Contributing

Contributions welcome! Please open an issue or PR.

---

**Current Version**: 0.1.0  
**Status**: Phases 1-3 Complete ✅  
**Updated**: 2026-02-16
