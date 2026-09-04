# SharePoint File Picker — Kontent.ai Custom Element

A [Kontent.ai custom element](https://kontent.ai/learn/docs/custom-elements) that lets content
editors pick one or more files from SharePoint / OneDrive (via the Microsoft
[File Picker v8](https://learn.microsoft.com/en-us/onedrive/developer/controls/file-pickers/js-v8/open-file?view=odsp-graph-online))
and stores the selected files' name, URL, author and last-modified date as the element's value.

No server component is required — it's a static page hosted anywhere over HTTPS. Authentication
happens client-side against Azure AD using [MSAL.js](https://github.com/AzureAD/microsoft-authentication-library-for-js).

## How it works

1. The editor clicks **Select SharePoint Files**.
2. The element signs the editor in with Azure AD (MSAL, popup flow) and silently acquires tokens
   for OneDrive/SharePoint and Microsoft Graph.
3. It opens the Microsoft File Picker in a popup window and relays authentication requests from
   the picker to MSAL over the picker's messaging channel.
4. Once files are selected, it enriches them with author/last-modified metadata via Microsoft
   Graph and saves the result as JSON on the content item.

## 1. Register an Azure AD application

You need an Azure AD app registration that the element will use to sign editors in.

1. In the [Azure Portal](https://portal.azure.com), go to **Azure Active Directory → App
   registrations → New registration**.
2. Set **Redirect URI** (platform: Single-page application) to the exact URL where you will host
   `index.html`, e.g. `https://your-host.example.com/index.html`.
3. Under **API permissions**, add:
   - **Microsoft Graph** (delegated) — `Files.Read.All` and `Sites.Read.All`, used for the
     author/last-modified metadata lookup after a file is picked.
   - **SharePoint** (delegated) — `AllSites.Read` and `MyFiles.Read`. The File Picker itself
     authenticates directly against your tenant's SharePoint resource
     (`https://<tenant>.sharepoint.com` and `https://<tenant>-my.sharepoint.com`), which is a
     separate API from Graph — Graph permissions alone are not enough. If "SharePoint" isn't in
     the API picker's default list, search for it under "APIs my organization uses".

   Grant admin consent for your tenant after adding both sets of permissions.
4. Copy the **Application (client) ID** — you'll need it for the element configuration below.

Consult Microsoft's [File Picker v8 authentication docs](https://learn.microsoft.com/en-us/onedrive/developer/controls/file-pickers/js-v8/authentication?view=odsp-graph-online)
for the exact permission set your tenant/picker configuration requires.

## 2. Host the files

This is a static site — any HTTPS static host works (GitHub Pages, Azure Static Web Apps, Netlify,
an S3/Blob Storage bucket behind a CDN, etc.). Host the whole repository root (`index.html`,
`css/`, `js/`) and note the public URL of `index.html`; that's your **Hosted code URL** for step 3.

The redirect URI registered in Azure AD (step 1) must match this URL exactly.

For local development:

```bash
npm start
```

This serves the project at `http://localhost:5500`. Note that Azure AD app registrations require
`http://localhost` (not HTTPS) to be listed as an additional redirect URI if you want to test sign-in
locally.

## 3. Add the custom element in Kontent.ai

1. In your content type, add a **Custom element**.
2. Set **Hosted code URL** to the public URL of `index.html` from step 2.
3. In **Custom element configuration (JSON)**, provide:

```json
{
  "clientId": "00000000-0000-0000-0000-000000000000",
  "sharePointTenant": "contoso",
  "selectionMode": "multiple",
  "debug": false
}
```

| Field              | Required | Description                                                                                   |
|--------------------|----------|-------------------------------------------------------------------------------------------------|
| `clientId`         | Yes      | The Azure AD application (client) ID from step 1.                                              |
| `sharePointTenant` | Yes      | Your tenant's short name, i.e. the `contoso` in `contoso.sharepoint.com`.                        |
| `selectionMode`    | No       | `"multiple"` (default) or `"single"`.                                                          |
| `debug`            | No       | `true` to show a small status line useful for troubleshooting auth/picker issues. Default `false`. |

The element validates this configuration on load and shows an inline error if it's missing or
malformed, instead of failing silently.

> This setup assumes your tenant uses the default SharePoint domains
> (`contoso.sharepoint.com` / `contoso-my.sharepoint.com`). Tenants on a custom/vanity domain
> aren't supported without code changes.

## Stored value

The element stores a JSON array on the content item, e.g.:

```json
[
  {
    "id": "01ABCDEF...",
    "driveId": "b!AbCdEf...",
    "name": "Report.pdf",
    "url": "https://contoso-my.sharepoint.com/personal/.../Report.pdf",
    "author": "Jane Doe",
    "lastModified": "2026-01-15T10:30:00Z"
  }
]
```

## Notes for integrators

- **Embedding required**: Kontent.ai always loads custom elements inside a child iframe. If
  `index.html` is opened directly (`window === window.top`), the element shows a plain message
  instead of the picker UI rather than trying to run standalone.
- **Popups**: sign-in and the file picker itself both use popup windows. Editors must allow popups
  for the page the custom element is hosted on.
- **No `alert()`/`confirm()`**: Kontent.ai renders custom elements inside a sandboxed iframe that
  does not support native modal dialogs, so all error and confirmation UI in this element is
  implemented inline (banners) rather than with `alert()`/`confirm()`.
- **Origin checks**: messages from the file picker popup are only accepted if they come from
  `https://<tenant>.sharepoint.com` or `https://<tenant>-my.sharepoint.com`, matched by exact
  origin (not substring), to prevent a malicious page from spoofing the picker's messages.
- **Content-Security-Policy**: `index.html` ships a restrictive CSP meta tag. If you fork this and
  add other script/style sources, update it accordingly.

## License

MIT — see [LICENSE](LICENSE).
