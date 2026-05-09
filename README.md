# MSContactsSync

A Windows 10 Mobile (W10M) UWP app that synchronizes Microsoft Account contacts (Hotmail / Outlook.com) to the phone's People hub using the Microsoft Graph API.

## Authentication

The app uses standard **OAuth2 Authorization Code Flow** — the same sign-in experience you get on any modern website or app.

Here is how it works:

1. Tap **Sign in with Microsoft**
2. The Microsoft sign-in page opens in the Edge browser
3. Enter your Microsoft account email
4. Enter your password
5. Review and accept the permissions the app is requesting
6. The browser closes automatically and the app is ready to sync

Your credentials are never entered into the app itself. Everything happens through Microsoft's own sign-in page in the Edge browser, exactly as you would sign in to Outlook.com or any other Microsoft service.

## Privacy and Security

- The app never sees your Microsoft account password
- All authentication is handled entirely by Microsoft's servers
- The app only receives an access token — a temporary key that grants read access to your contacts
- The token is stored locally on your phone in the app's private storage
- No data is sent anywhere except to Microsoft Graph API (`graph.microsoft.com`)
- The app is read-only — it only downloads contacts to the phone and does not modify your Microsoft account contacts

## Client ID

The app includes a built-in **App Registration Client ID** — an identifier registered with Microsoft Azure that allows the app to request access to the Contacts API on behalf of the user.

**This Client ID is not a secret.** It identifies the application, not the user. Anyone who has it can only use it to trigger the standard Microsoft sign-in flow — they cannot access any user's data without that user explicitly signing in and granting permission.

The built-in Client ID is obfuscated in the source code (XOR encoding) simply to avoid it being indexed by code search engines — not for security reasons.

## Using Your Own Client ID

If you prefer to register your own Azure application instead of using the built-in one:

1. Go to https://portal.azure.com
2. Search for **App registrations** → **New registration**
3. Name: anything you like
4. Supported account types: **Personal Microsoft accounts only**
5. Click **Register**
6. Go to **Authentication** → **Add a platform** → **Mobile and desktop applications**
7. Check `https://login.microsoftonline.com/common/oauth2/nativeclient` → **Configure**
8. Go to **Settings** tab → enable **Allow public client flows** → **Save**
9. Go to **API permissions** → **Add** → **Microsoft Graph** → **Delegated:**
   - `User.Read`
   - `Contacts.Read`
10. Copy the **Application (client) ID** from the **Overview** page

In the app, check **"Use custom App Registration ID"** on the Sign in screen and paste your Client ID there.

To embed your own ID in the source code instead, replace `_clientIdBytes` in `MainPage.xaml.cs`:

```python
# Run this Python one-liner to generate the obfuscated bytes for your Client ID
python3 -c "print([hex(ord(c)^42) for c in 'your-client-id-here'])"
```

## How Sync Works

- **First sync** — downloads all contacts from Microsoft and saves them to the People hub
- **Subsequent syncs** — uses the Microsoft Graph **delta query** API, which returns only contacts that changed since the last sync; unchanged contacts are skipped entirely
- **Deleted contacts** — contacts deleted from your Microsoft account are removed from the phone on the next sync

## Field Mapping

| Microsoft Graph        | People Hub              |
|------------------------|-------------------------|
| givenName              | First name              |
| middleName             | Middle name             |
| surname                | Last name               |
| companyName            | Company                 |
| jobTitle               | Job title               |
| mobilePhone            | Mobile phone            |
| businessPhones         | Work phones             |
| homePhones             | Home phones             |
| emailAddresses         | Email addresses         |
| businessAddress        | Work address            |
| homeAddress            | Home address            |
| otherAddress           | Other address           |

## Building

- Visual Studio 2017
- Windows 10 SDK 10.0.15063.0
- Target: Windows 10 Mobile (ARM)
- Build Release|ARM and deploy directly

## License

Open source. Use freely. Contributions welcome.