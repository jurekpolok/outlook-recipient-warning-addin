# Outlook Recipient Privacy Warning Add-in

An Outlook Add-in that warns users when they add more than 5 recipients in the To and CC fields, recommending the use of BCC for external recipients to protect privacy.

## Features

- **Automatic monitoring**: Detects when recipients are added to To/CC fields
- **Threshold warning**: Shows a warning when more than 5 recipients are in To/CC
- **External recipient detection**: Identifies recipients from outside your organization
- **Privacy guidance**: Recommends using BCC for external recipients
- **Task pane UI**: Visual interface showing recipient counts and warnings

## Configuration

### Set Your Internal Domains

Edit the `INTERNAL_DOMAINS` array in both `src/taskpane.js` and `src/functions.js`:

```javascript
const INTERNAL_DOMAINS = [
    "yourcompany.com",
    "yourcompany.org"
];
```

### Adjust the Threshold

Change `RECIPIENT_THRESHOLD` in both files (default is 5):

```javascript
const RECIPIENT_THRESHOLD = 5;
```

## Installation

### Prerequisites

- Node.js (v14 or higher)
- npm
- Outlook (desktop, web, or mobile)

### Development Setup

1. **Install dependencies**:
   ```bash
   cd outlook-recipient-warning-addin
   npm install
   ```

2. **Generate SSL certificates** (required for local development):
   ```bash
   # Using mkcert (recommended)
   brew install mkcert  # macOS
   mkcert -install
   mkcert localhost

   # This creates localhost.pem and localhost-key.pem
   # Rename them:
   mv localhost.pem localhost.crt
   mv localhost-key.pem localhost.key
   ```

3. **Start the local server**:
   ```bash
   npm start
   ```
   The add-in will be served at `https://localhost:3000`

4. **Sideload the add-in**:

   **For Outlook on the web:**
   - Go to Outlook.com or your Microsoft 365 Outlook
   - Click the gear icon > View all Outlook settings
   - Go to Mail > Customize actions > Add-ins
   - Click "My add-ins" > "Add a custom add-in" > "Add from file"
   - Upload the `manifest.xml` file

   **For Outlook desktop (Windows):**
   - Open Outlook
   - Go to File > Manage Add-ins (or Get Add-ins)
   - Click "My add-ins" > "Add a custom add-in" > "Add from file"
   - Select the `manifest.xml` file

   **For Outlook desktop (Mac):**
   - Open Outlook
   - Go to Tools > Get Add-ins
   - Click "My add-ins" > "Add a custom add-in" > "Add from file"
   - Select the `manifest.xml` file

## Production Deployment

For production use:

1. **Host the files** on a web server with HTTPS (required)
2. **Update URLs** in `manifest.xml` to point to your production server
3. **Generate a unique GUID** for the `<Id>` element in manifest.xml
4. **Deploy via Microsoft 365 Admin Center** for organization-wide deployment:
   - Go to admin.microsoft.com
   - Navigate to Settings > Integrated apps
   - Click "Upload custom apps"
   - Upload your manifest.xml

## File Structure

```
outlook-recipient-warning-addin/
├── manifest.xml          # Add-in manifest configuration
├── package.json          # Node.js dependencies
├── README.md             # This file
└── src/
    ├── taskpane.html     # Task pane UI
    ├── taskpane.js       # Task pane logic
    ├── taskpane.css      # Task pane styles
    ├── functions.html    # Event handler host page
    └── functions.js      # Event-based automation
```

## How It Works

1. **Task Pane**: Users can click "Check Recipients" button to open a panel showing:
   - Current recipient counts (To, CC, BCC)
   - Warning if threshold is exceeded
   - Count of external recipients

2. **Event-based**: The add-in also monitors recipient changes automatically (when supported by the Outlook client) and shows an info bar notification.

## Requirements

- Office.js API version 1.3 or higher
- Mailbox requirement set 1.3 or higher

## Troubleshooting

- **Add-in not loading**: Ensure your server is running with HTTPS
- **No warnings appearing**: Check that you've configured internal domains correctly
- **Event handler not triggering**: Event-based activation requires newer Outlook versions

## License

MIT
