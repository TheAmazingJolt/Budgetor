# Budget Automator — Integration Setup Guide

Complete step-by-step instructions for setting up all integrations on your Railway deployment.

**Railway API Server URL:**
```
https://workspaceapi-server-production-01e3.up.railway.app
```

> Replace this URL throughout this guide if your Railway deployment URL is different.

---

## Prerequisite: Session Secret

The API server requires a session secret in production. Without it, the server will refuse to start.

1. Generate a random string:
   ```
   openssl rand -hex 32
   ```
2. In Railway, add the following environment variable to your API server service:

| Variable | Value |
|---|---|
| `SESSION_SECRET` | The random string you just generated |

---

## 1. Google Sheets Integration (Read/Write Spreadsheets)

This lets users connect their Google account so the app can directly read their existing budget sheets and write new weekly columns back.

### Step 1 — Create a Google Cloud Project

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project (or use an existing one)
3. Go to **APIs & Services > Library** and enable:
   - **Google Sheets API**
   - **Google Drive API**

### Step 2 — Create OAuth Credentials

1. Go to **APIs & Services > Credentials > Create Credentials > OAuth Client ID**
2. Application type: **Web application**
3. Add this to **Authorized redirect URIs:**
   ```
   https://workspaceapi-server-production-01e3.up.railway.app/api/auth/google/callback
   ```
4. Copy the **Client ID** and **Client Secret**

### Step 3 — Set Railway Environment Variables

| Variable | Value |
|---|---|
| `GOOGLE_CLIENT_ID` | Client ID from Step 2 |
| `GOOGLE_CLIENT_SECRET` | Client Secret from Step 2 |
| `GOOGLE_REDIRECT_URI` | `https://workspaceapi-server-production-01e3.up.railway.app/api/auth/google/callback` |

### Optional — Public Sheet Reading Without Sign-In

If you want users to paste a public Google Sheets URL and have it work without signing in first:

1. Go to **APIs & Services > Credentials > Create Credentials > API Key**
2. Restrict the key to the **Google Sheets API** only
3. Add to Railway:

| Variable | Value |
|---|---|
| `GOOGLE_API_KEY` | The API key you created |

---

## 2. Google Sign-In (User Accounts)

This lets users sign in to their Budget Automator account using their Google profile. This is separate from the Sheets connection above, but uses the **same Google Cloud project and same Client ID/Secret**.

### Step 1 — Add a Second Redirect URI

Go back to the same OAuth Client you created in Section 1, Step 2. Add a **second** Authorized redirect URI:

```
https://workspaceapi-server-production-01e3.up.railway.app/api/auth/login/google/callback
```

### Step 2 — Set Railway Environment Variable

| Variable | Value |
|---|---|
| `GOOGLE_ACCOUNT_REDIRECT_URI` | `https://workspaceapi-server-production-01e3.up.railway.app/api/auth/login/google/callback` |

No new Client ID or Client Secret is needed — the same `GOOGLE_CLIENT_ID` and `GOOGLE_CLIENT_SECRET` from Section 1 are reused.

---

## 3. Microsoft Excel Online Integration (Read/Write OneDrive Files)

This lets users connect their Microsoft account so the app can list their OneDrive Excel files, read budget data from them, and write new weeks back.

### Step 1 — Register an Azure App

1. Go to [portal.azure.com](https://portal.azure.com)
2. Navigate to **Azure Active Directory > App registrations > New registration**
3. Fill in:
   - **Name**: anything (e.g. "Budget Automator")
   - **Supported account types**: Select **"Accounts in any organizational directory and personal Microsoft accounts"** (the Multitenant + personal option — this covers Outlook.com, Hotmail, and work/school accounts)
   - **Redirect URI**: Select **Web**, then enter:
     ```
     https://workspaceapi-server-production-01e3.up.railway.app/api/auth/microsoft/callback
     ```
4. Click **Register**

### Step 2 — Create a Client Secret

1. In your app registration, go to **Certificates & secrets > New client secret**
2. Set an expiry (24 months is fine)
3. **Copy the Value immediately** — you cannot see it again after leaving the page

### Step 3 — Add API Permissions

1. Go to **API permissions > Add a permission > Microsoft Graph > Delegated permissions**
2. Add these permissions:
   - `Files.ReadWrite`
   - `User.Read`
   - `openid`
   - `offline_access`
3. Click **Grant admin consent** if prompted (not always required for personal accounts)

### Step 4 — Set Railway Environment Variables

| Variable | Value |
|---|---|
| `MICROSOFT_CLIENT_ID` | The **Application (client) ID** from the app's Overview page |
| `MICROSOFT_CLIENT_SECRET` | The secret **Value** from Step 2 |
| `MICROSOFT_REDIRECT_URI` | `https://workspaceapi-server-production-01e3.up.railway.app/api/auth/microsoft/callback` |

---

## 4. Apple Sign-In (User Accounts)

This lets users sign in to their Budget Automator account with their Apple ID. This is the most involved setup and requires an Apple Developer account ($99/year).

### Step 1 — Configure Your App ID

1. Go to [developer.apple.com](https://developer.apple.com) > **Certificates, IDs & Profiles > Identifiers**
2. Create or select an App ID
3. Enable the **Sign In with Apple** capability

### Step 2 — Create a Services ID

1. Go to **Identifiers > + > Services IDs**
2. Enter a description and a unique identifier (e.g. `com.yourcompany.budgetautomator.web`) — this becomes your `APPLE_CLIENT_ID`
3. Enable **Sign In with Apple**, then click **Configure**:
   - **Primary App ID**: Select the App ID from Step 1
   - **Domains**: Your Railway frontend domain (e.g. `budget-automator.up.railway.app`)
   - **Return URL**:
     ```
     https://workspaceapi-server-production-01e3.up.railway.app/api/auth/login/apple/callback
     ```

### Step 3 — Create a Private Key

1. Go to **Keys > + > Sign In with Apple**, select your App ID
2. Download the `.p8` file (only downloadable once — keep it safe)
3. Note the **Key ID** shown on the page

### Step 4 — Set Railway Environment Variables

| Variable | Value |
|---|---|
| `APPLE_CLIENT_ID` | The Services ID identifier from Step 2 (e.g. `com.yourcompany.budgetautomator.web`) |
| `APPLE_REDIRECT_URI` | `https://workspaceapi-server-production-01e3.up.railway.app/api/auth/login/apple/callback` |
| `APPLE_TEAM_ID` | Your 10-character Team ID (visible in the top right on developer.apple.com) |
| `APPLE_KEY_ID` | The Key ID from Step 3 |
| `APPLE_PRIVATE_KEY` | The entire contents of the `.p8` file, with actual newlines replaced by `\n` so it fits on one line |

---

## 5. CORS Configuration

If your frontend and API server are deployed as separate Railway services (which is the typical setup), add this to the API server:

| Variable | Value |
|---|---|
| `CORS_ORIGIN` | Your frontend's Railway URL (e.g. `https://budget-automator-frontend.up.railway.app`). Separate multiple origins with commas. |

---

## Complete Environment Variable Reference

| Variable | Required For | Notes |
|---|---|---|
| `SESSION_SECRET` | Everything | Set this first; server won't start without it in production |
| `DATABASE_URL` | Everything | Auto-set by Railway if you add a Postgres plugin |
| `GOOGLE_CLIENT_ID` | Google Sheets + Google Sign-In | One Client ID for both features |
| `GOOGLE_CLIENT_SECRET` | Google Sheets + Google Sign-In | One Client Secret for both features |
| `GOOGLE_REDIRECT_URI` | Google Sheets OAuth | Callback URL for spreadsheet read/write |
| `GOOGLE_ACCOUNT_REDIRECT_URI` | Google Sign-In | Callback URL for user account login |
| `GOOGLE_API_KEY` | Optional | Enables public sheet URL paste without sign-in |
| `MICROSOFT_CLIENT_ID` | Microsoft Excel / OneDrive | Application (client) ID from Azure |
| `MICROSOFT_CLIENT_SECRET` | Microsoft Excel / OneDrive | Client secret value from Azure |
| `MICROSOFT_REDIRECT_URI` | Microsoft Excel / OneDrive | Callback URL for OneDrive access |
| `APPLE_CLIENT_ID` | Apple Sign-In | Services ID identifier |
| `APPLE_REDIRECT_URI` | Apple Sign-In | Callback URL for Apple login |
| `APPLE_TEAM_ID` | Apple Sign-In | 10-character Team ID |
| `APPLE_KEY_ID` | Apple Sign-In | Key ID for the Sign-In key |
| `APPLE_PRIVATE_KEY` | Apple Sign-In | Contents of `.p8` file (newlines as `\n`) |
| `CORS_ORIGIN` | Cross-origin deployment | Frontend URL(s), comma-separated |

---

Each feature works independently. The app automatically hides buttons and cards for providers that aren't configured, so nothing will break if you skip a section.
