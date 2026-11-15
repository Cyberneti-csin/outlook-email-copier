# Outlook Email Copier

A standalone, static browser app for browsing Outlook messages and copying selected emails as plain text. It uses Microsoft Graph delegated access and does not have a server.

## Publish on GitHub Pages

This app is prepared to deploy from its own public GitHub repository. Do not publish the parent Brewery Quality Standard repository if you only intend to share this app.

1. On GitHub, create a **new public repository** named `outlook-email-copier`. Do not add a README, license, or starter workflow.
2. In the new repository, open **Settings → Pages** and set the build and deployment source to **GitHub Actions** before pushing code.
3. In PowerShell, run these commands to publish only this standalone app folder. Replace `YOUR-USERNAME` with your GitHub username:

   ```powershell
   Set-Location C:\Users\Administrator\outlook-email-copier
   git init -b main
   git add .
   git commit -m "Initial Outlook Email Copier app"
   git remote add origin https://github.com/YOUR-USERNAME/outlook-email-copier.git
   git push -u origin main
   ```

   This pushes the app files, including `.github/workflows/deploy.yml`; it does not push the parent Brewery Quality Standard repository.
4. Open **Actions** and wait for **Deploy Outlook Email Copier** to finish successfully. The Pages site address is shown in the deployment details and normally looks like `https://GITHUB-USERNAME.github.io/outlook-email-copier/`.
5. In Microsoft Entra, open this app registration's **Authentication** settings and add that exact HTTPS site address as a **Single-page application** redirect URI. Keep `http://localhost:8000/` if you also want local testing.
6. Open the published site and sign in with Microsoft. Each person must use their own account and grant the app the delegated `Mail.Read` permission.

The public repository makes the app's code available to everyone; it does not publish email or provide access to your inbox. The app ID is public by design, but never publish a client secret or access token. Microsoft may show a consent warning for an unverified publisher; users should only grant access if they trust the app and understand its mail-reading permission.

## Set up Microsoft sign-in

The app needs an app registration in your Microsoft account. Registration is free. The app ID is not a password; it identifies the registration.

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com/) and create a new app registration.
2. Choose the account type that matches your Outlook account. To support both personal Outlook.com accounts and work or school accounts, choose **Accounts in any organizational directory and personal Microsoft accounts**.
3. Under **Authentication**, add a **Single-page application (SPA)** platform and enter the app's exact URL as its redirect URI. For local testing, serve this folder on `http://localhost:8000` and enter `http://localhost:8000/` as the redirect URI. For a deployed site, use the full URL of this app's `index.html`.
4. Under **API permissions**, add the Microsoft Graph **delegated** permission `Mail.Read`. The app does not need write or send-mail permissions. Your organization may require an administrator to approve this permission.
5. Copy the **Application (client) ID** from the app registration and enter it into the app's connection form. If the app ID has been preconfigured in your local copy, it will already be filled in.

Do not open `index.html` directly as a `file://` URL. Microsoft sign-in requires the app to be served from a registered web address. For example, from this folder run:

```powershell
python -m http.server 8000
```

Then open `http://localhost:8000/` in your browser.

## Use

Connect your account, choose a mail folder, and select the emails you want. Search filters the messages loaded from the current folder; use **Load more** to fetch additional messages. **Copy selected** puts a plain-text copy of each selected email on your clipboard, including its subject, sender, received date, and message body.

## Privacy

Email is requested directly from Microsoft Graph in your browser. The app has no backend and does not upload, store, or send your messages. The client ID is saved in this browser's local storage; sign-in tokens are held in session storage. Selecting **Sign out** clears the app's in-memory email data. Copying is an explicit action, so review the selected messages before pasting their content into Copilot.

The app requests only the delegated `Mail.Read` permission. Microsoft may display its normal consent and sign-in screens.
