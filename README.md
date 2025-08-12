# TrueLayer to Google Sheets

A Google Apps Script project that automatically imports banking data from TrueLayer into Google Sheets. This project connects to your bank accounts via TrueLayer's Open Banking API and automatically syncs transactions, accounts, and card data to organized Google Sheets.

## Features

- 🏦 Connect multiple bank accounts and credit cards
- 🔄 Automatic hourly transaction synchronization
- 📊 Organized data in separate sheets (Accounts, yearly transaction sheets)
- 🔐 Secure OAuth authentication via TrueLayer
- 📱 QR code authentication for mobile convenience
- 💾 Smart duplicate detection and data management

## Prerequisites

1. **Google Account** with access to Google Sheets and Google Apps Script
2. **TrueLayer Developer Account** - Sign up at [TrueLayer Console](https://console.truelayer.com/)
3. **Google Cloud Platform Project** (for advanced features and production use)

## Setup Instructions

### 1. Create a Google Cloud Platform Project

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Click **"Create Project"** or select an existing project
3. Choose a project name (e.g., "TrueLayer-Sheets-Integration")
4. Note down your **Project ID** (you'll need this later)
5. Enable the required APIs:
   - Go to **APIs & Services > Library**
   - Search for and enable:
     - **Google Apps Script API**
     - **Google Sheets API**
     - **Google Drive API**

### 2. Set up TrueLayer Application

1. Go to [TrueLayer Console](https://console.truelayer.com/)
2. Create a new application or use an existing one
3. Configure your application settings:
   - **Environment**: Choose Sandbox (for testing) or Live (for production)
   - **Redirect URIs**: You'll add this after deploying the script (Step 4)
   - **Scopes**: Ensure you have the following permissions:
     - `info`
     - `accounts`
     - `transactions`
     - `balance`
     - `cards`
     - `offline_access`
4. Note down your **Client ID** and **Client Secret**

### 3. Create Google Sheets and Deploy the Script

1. **Create a new Google Spreadsheet**:
   - Go to [Google Sheets](https://sheets.google.com)
   - Create a new blank spreadsheet
   - Give it a meaningful name (e.g., "My Banking Data")

2. **Open Google Apps Script**:
   - In your spreadsheet, go to **Extensions > Apps Script**
   - This will create a new Apps Script project bound to your spreadsheet

3. **Replace the default code**:
   - Delete the default `myFunction()` code
   - Copy and paste the contents of `Code.gs` from this repository
   - Create additional files by clicking the **"+"** next to "Files":
     - `settings.html` - Copy content from the repository
     - `connectDialog.html` - Copy content from the repository
   - **Enable and update `appsscript.json`**:
     - In the Apps Script editor, click on **Project Settings** (gear icon) in the left sidebar
     - Check the box **"Show 'appsscript.json' manifest file in editor"**
     - Go back to the **Editor** tab - you should now see `appsscript.json` in the file list
     - Replace the contents of `appsscript.json` with the content from the repository

4. **Set up Google Cloud Project (Optional but Recommended)**:
   - In Apps Script, go to **Project Settings** (gear icon)
   - Scroll down to **Google Cloud Platform (GCP) Project**
   - Click **"Change project"**
   - Enter your GCP Project ID from Step 1
   - Click **"Set project"**

### 4. Deploy as Web App

1. **Deploy the script**:
   - Click **Deploy > New Deployment**
   - Click the gear icon next to "Type" and select **"Web app"**
   - Configure deployment settings:
     - **Description**: "TrueLayer Banking Integration"
     - **Execute as**: "Me"
     - **Who has access**: "Anyone with the link"
   - Click **"Deploy"**
   - **Authorize** the permissions when prompted
   - Copy the **Web app URL** (you'll need this for TrueLayer)

2. **Update TrueLayer Redirect URI**:
   - Go back to [TrueLayer Console](https://console.truelayer.com/)
   - Open your application settings
   - Add the Web app URL from step 1 to **Redirect URIs**
   - Save the changes

### 5. Configure the Integration

1. **Open your Google Spreadsheet**
2. **Access the menu**:
   - You should see a new menu called **"Bank Connector"**
   - If you don't see it, refresh the page and wait a moment
3. **Configure settings**:
   - Click **Bank Connector > Settings**
   - Enter your TrueLayer **Client ID** and **Client Secret**
   - The **Redirect URI** should automatically show your web app URL
   - Click **"Save"**

### 6. Connect Your Bank

1. **Start the connection process**:
   - Click **Bank Connector > Connect Bank**
   - A dialog will appear with a QR code and link
2. **Authenticate**:
   - **Mobile**: Scan the QR code with your phone
   - **Desktop**: Click the link
3. **Select your bank** and complete the authentication process
4. **Authorization**: Complete the OAuth flow with your bank
5. **Confirmation**: You should see "Bank connected. You can close this window."

## Usage

### Manual Refresh
- Use **Bank Connector > Refresh** to manually sync the latest transactions

### Automatic Refresh
- The script automatically creates an hourly trigger to sync transactions
- Data is refreshed every hour without manual intervention

### Data Organization

The integration creates several sheets in your spreadsheet:

- **Accounts**: Contains all connected accounts and cards with metadata
- **YYYY** (e.g., "2024"): Annual sheets containing transactions for each year
- **_config** (hidden): Stores configuration and authentication tokens

### Data Fields

**Accounts Sheet**:
- Account ID, Display Name, Type, Currency
- Account numbers, IBAN, SWIFT BIC
- Provider information and connection status

**Transaction Sheets**:
- Date, Description, Amount, Currency
- Account ID, Category, External ID
- Full transaction details in Notes field

## Security Considerations

- 🔐 All authentication tokens are stored securely in your private spreadsheet
- 🚫 Never share your Client Secret or spreadsheet with untrusted parties
- ✅ Use the Sandbox environment for testing
- 🔄 Access tokens are automatically refreshed
- 📱 OAuth flow ensures you never share banking credentials

## Troubleshooting

### Common Issues

**"Please set Client ID and Secret first" error**:
- Ensure you've entered the correct TrueLayer credentials in Settings
- Check that you've copied them without extra spaces

**Web app URL not working**:
- Ensure the deployment is set to "Anyone with the link"
- Redeploy if you made changes to the script
- Update the redirect URI in TrueLayer Console

**No menu appearing**:
- Refresh the spreadsheet page
- Check that the script is properly saved and deployed
- Ensure you've copied all files correctly

**Authentication failures**:
- Verify your TrueLayer app has the correct scopes enabled
- Check that the redirect URI exactly matches your web app URL
- Ensure your TrueLayer app is in the correct environment (Sandbox/Live)

### Debug Mode

To enable detailed logging:
1. In Apps Script, go to **Execution > View executions**
2. Run functions manually to see detailed error messages
3. Check the **Logs** for debugging information

## Development

### Local Development
1. Use the Apps Script editor for development
2. Test functions individually using the built-in debugger
3. Use `Logger.log()` for debugging output

### Contributing
1. Fork the repository
2. Make your changes
3. Test thoroughly with Sandbox environment
4. Submit a pull request

## API Limits and Costs

- **Google Apps Script**: Free tier includes 6 minutes of execution time per trigger
- **TrueLayer**: Check your plan limits for API calls
- **Google Sheets**: No additional costs for data storage

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

Please ensure compliance with:
- TrueLayer's Terms of Service
- Google's Apps Script Terms of Service
- Your bank's terms and conditions
- Local data protection regulations (GDPR, etc.)

## Support

For issues with:
- **This integration**: Create an issue in this repository
- **TrueLayer API**: Contact TrueLayer support
- **Google Apps Script**: Check Google's documentation and community forums

---

**⚠️ Important**: This integration handles sensitive financial data. Always use strong authentication, keep your credentials secure, and regularly review your connected applications and permissions.
