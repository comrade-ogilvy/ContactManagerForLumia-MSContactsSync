# Create a Microsoft Entra App for Contacts Access

This guide explains how to create or configure a Microsoft Entra / Azure application for personal Microsoft accounts such as Outlook.com, Hotmail, Live, or MSN so a contacts application can access Microsoft Graph.

The app is intended for **contacts only**. It uses delegated Microsoft Graph permissions, meaning the app can access contacts only after you sign in and approve the requested permissions.

---

## 1. Open App Registrations

Go to:

```text
https://entra.microsoft.com/
```

Then open:

```text
Identity
→ Applications
→ App registrations
```

Click:

```text
New registration
```

---

## 2. Register the Application

Use an application name that describes contacts access, for example:

```text
Contacts Management
```

For **Supported account types**, choose the option that includes personal Microsoft accounts:

```text
Accounts in any organizational directory and personal Microsoft accounts
```

This is important for personal Microsoft accounts such as Outlook.com, Hotmail, Live, and MSN.

For **Redirect URI**, you can leave it empty during initial registration and configure it later under **Authentication**.

Click:

```text
Register
```

---

## 3. Save the Application Client ID

After registration, open the application **Overview** page.

Copy this value:

```text
Application (client) ID
```

You will need to enter this value into your contacts application.

Do **not** use the **Directory / Tenant ID** as the application client ID.
**This ID is used to authenticate the app (not your Microsoft account)**

---

## 4. Configure Authentication

Open:

```text
Manage
→ Authentication
```

Click:

```text
Add a platform
```

Choose:

```text
Mobile and desktop applications
```

Add or select an appropriate redirect URI for a mobile or desktop application.

Then save the changes.

The important part is that this app is treated as a **public client** app, not as a confidential web-server app.

If you see an option like this, enable it:

```text
Allow public client flows
```

or:

```text
Treat application as a public client
```

Set it to:

```text
Yes
```

Then click:

```text
Save
```

---

## 5. Add Microsoft Graph API Permissions

Open:

```text
Manage
→ API permissions
```

Click:

```text
Add a permission
```

Choose:

```text
Microsoft Graph
```

Choose:

```text
Delegated permissions
```

Delegated permissions are correct because the app signs in as you and accesses only the data you approve.

---

## 6. Contacts Permissions

For contacts access, add:

```text
User.Read
Contacts.Read
```

`Contacts.Read` is enough for read-only contacts access.

Do not add write permissions unless the application really needs to create, update, or delete contacts.

---

## 7. Do You Need Admin Consent?

For your own personal Microsoft account, usually no.

You may see a consent screen saying the app is:

```text
Unverified
```

That is expected for a personal, unpublished app.

This does **not** automatically mean the app is blocked.

You can usually continue and approve the requested permissions for your own account.

Admin consent and publisher verification are mostly important if you distribute the app to many other users or organizations.

---

## 8. Publisher Domain / Unverified App Warning

You may see something like:

```text
Publisher domain: something.onmicrosoft.com
The application’s consent screen will show “Unverified”.
onmicrosoft.com publisher domains are not allowed for publisher verification.
```

For a private contacts app, this is usually safe to ignore.

It means Microsoft has not verified you as a public software publisher.

You only need publisher verification if you want a verified badge or if you want many unrelated users to grant consent to your app.

For a private contacts app used by your own accounts, you do not need:

```text
Custom domain
Partner Center account
Publisher verification
```

---

## 9. Rename the Application

To rename the app later, open:

```text
Microsoft Entra admin center
→ Identity
→ Applications
→ App registrations
→ All applications
→ select your app
→ Manage
→ Branding & properties
→ Name
→ Save
```

Renaming the app does **not** change the application client ID.

---

## 10. First Sign-In Behavior

On first sign-in, the application may ask you to authenticate and approve requested permissions.

After successful authentication, future sign-ins may not require approval again unless:

- The saved sign-in state is removed
- New permissions are added
- Microsoft requires re-consent
- You sign in with a different account

---

## 11. Troubleshooting

### Error: invalid_client

Check that the application client ID is correct.

Also make sure the app supports personal Microsoft accounts.

---

### Error: invalid_scope

Check that the permission exists and was added as a Microsoft Graph delegated permission.

For contacts access, use:

```text
Contacts.Read
```

---

### Error: interaction_required

Sign in again and approve the requested contacts permission.

---

### Error: Need admin approval

This usually happens with organizational Microsoft 365 accounts, not normal personal Outlook.com or Hotmail accounts.

For personal Microsoft accounts, make sure the app registration supports personal Microsoft accounts.

---

### Consent screen says Unverified

For a private app, this is expected.

Continue only if the app is yours and the application client ID matches your app registration.

---

## 12. Microsoft Graph Endpoints Used

Contacts applications commonly use endpoints such as:

```text
/me/contacts
/me/contactFolders
/me/contactFolders/{id}/contacts
/me/contactFolders/{id}/childFolders
```

---

## 13. References

Microsoft Graph permissions and API access are configured through Microsoft Entra app registrations and Microsoft Graph delegated permissions.

Microsoft Graph contacts APIs support delegated `Contacts.Read` for reading contacts.

Useful Microsoft documentation:

```text
https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-configure-app-access-web-apis
https://learn.microsoft.com/en-us/entra/identity-platform/msal-client-applications
https://learn.microsoft.com/en-us/entra/identity-platform/msal-authentication-flows
https://learn.microsoft.com/en-us/graph/api/user-list-contacts
https://learn.microsoft.com/en-us/graph/api/contactfolder-list-contacts
```
