# 📧 Outlook Skill for OpenClaw

A skill to read, search, send, and manage Outlook emails and calendar events via the Microsoft Graph API.

---

## 📋 Prerequisites

- [OpenClaw](https://openclaw.ai) installed and running
- A **Microsoft Azure account**
- An **Outlook / Microsoft 365 account**

---

## 🚀 Installation

Install the skill via ClawHub:

```bash
npx clawhub install byungkyu/outlook-api --yes
```

Verify it's ready:

```bash
openclaw skills list
```

You should see `✓ ready │ 📦 outlook` in the list.

---

## 🔐 Azure App Setup

### Step 1 — Register an App in Azure Portal

1. Go to [portal.azure.com](https://portal.azure.com)
2. Search for **App registrations** → click **New registration**
3. Name it (e.g. `OpenClaw Outlook`)
4. Under **Supported account types**, select:
   - `Personal accounts only` → use `/consumers` endpoint
   - `Any Entra ID Tenant + Personal Microsoft accounts` → use `/common` endpoint
5. Set **Redirect URI** to: `http://localhost:3000/oauth/callback`
6. Click **Register**

### Step 2 — Copy Your Credentials

From the app overview page, copy:
- **Application (client) ID**
- **Directory (tenant) ID**

### Step 3 — Create a Client Secret

1. Go to **Certificates & secrets** → **New client secret**
2. Add a description and expiry date → click **Add**
3. ⚠️ **Immediately copy the `Value` column** — it disappears after you leave the page!
   - `Value` = the secret you need ✅
   - `Secret ID` = NOT what you need ❌

### Step 4 — Set API Permissions

1. Go to **API Permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**
2. Add these permissions:
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Calendars.ReadWrite`
   - `User.Read`
   - `offline_access`
3. Click **Grant admin consent**

---

## ⚙️ Configuration in OpenClaw

Paste this prompt into your OpenClaw chat (replace placeholders with your actual values):

```
Set up my Outlook skill with these credentials:
- Client ID: YOUR_CLIENT_ID
- Client Secret: YOUR_CLIENT_SECRET (the Value, not the Secret ID)
- Tenant ID: YOUR_TENANT_ID
- Redirect URI: http://localhost:3000/oauth/callback

Use the /consumers endpoint for personal Microsoft accounts.
Then run the OAuth authentication flow and give me the authorization link.
```

---

## 🔑 OAuth Authentication

After pasting the prompt above:

1. OpenClaw will generate an **authorization URL** — open it in your browser
2. Sign into your **Microsoft account** and approve the permissions
3. You'll be redirected to `http://localhost:3000/oauth/callback?code=...`
4. **Copy the full redirect URL** (even if the page doesn't load)
5. Paste it back into OpenClaw to complete the token exchange

> ⚠️ If you get error `AADSTS9002346`, make sure the authorization URL uses `/consumers` not your tenant ID.

> ⚠️ If you get error `AADSTS7000215`, you used the **Secret ID** instead of the **Secret Value** — go back and copy the correct one.

---

## 📬 Usage — Example Prompts

### Emails

```
Show my last 10 emails in Outlook inbox
```
```
Show all unread emails
```
```
Search emails from john@example.com
```
```
Send an email to john@example.com with subject "Hello" and body "Testing OpenClaw!"
```
```
Reply to the latest email in my inbox
```
```
Show emails with attachments from this week
```
```
Search my inbox for emails containing "invoice"
```
```
How many unread emails do I have?
```
```
Give me a summary of today's emails
```

### Calendar

```
Show my calendar events for today
```
```
What meetings do I have this week?
```
```
Create a meeting tomorrow at 3pm called "Team Sync" for 1 hour
```
```
Delete the event called "Team Sync" from my calendar
```

---

## 🛠️ Troubleshooting

| Error | Cause | Fix |
|-------|-------|-----|
| `Authentication failed` | Repo is private | Install via `npx clawhub install` not git clone |
| `AADSTS9002346` | Wrong endpoint (tenant vs consumers) | Change URL to use `/consumers` endpoint |
| `AADSTS7000215` | Used Secret ID instead of Secret Value | Copy the `Value` column from Azure portal |
| `No access token` | OAuth not completed | Re-run the authentication flow |
| `Exit code 1` | CLI not installed or bad credentials | Ask OpenClaw to show full error output |

---

## 🔄 Re-authenticate

If your token expires, paste this into OpenClaw:

```
Re-run the Outlook OAuth authentication flow from scratch and give me a new authorization link.
```

---

## 📁 Skill Info

| Field | Value |
|-------|-------|
| Skill Name | `outlook` |
| Source | `byungkyu/outlook-api` |
| Version | `1.0.3` |
| Registry | [clawhub.ai/byungkyu/outlook-api](https://clawhub.ai/byungkyu/outlook-api) |
| API | Microsoft Graph API |
| Auth | OAuth 2.0 (Authorization Code Flow) |

---

## 📄 License

See [ClawHub](https://clawhub.ai) for skill licensing terms.
