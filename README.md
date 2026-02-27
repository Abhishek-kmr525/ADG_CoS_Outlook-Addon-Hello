# Outlook Thank You Test Add-in

This folder contains a minimal Outlook add-in for organization-wide Microsoft 365 deployment.

## What it does
- Adds a ribbon button in Outlook for message read and compose.
- When clicked, it opens a popup dialog with `Thank You`.
- If a dialog cannot open, it shows an informational fallback message.

## Files
- `manifest.xml` - Upload this to Microsoft 365 Admin Center.
- `src/commands.html` and `src/commands.js` - Command runtime code.
- `src/thankyou-dialog.html` - Popup dialog page.
- `src/fallback.html` - Fallback page URL for legacy form settings.

## Before upload (required)
1. Host the `src` files and your icon files on an HTTPS domain reachable by Outlook clients.
2. In `manifest.xml`, replace all `https://YOUR_ADDIN_DOMAIN/...` values with your real HTTPS URLs.
3. Ensure these URLs are valid:
   - `.../src/commands.html`
   - `.../src/thankyou-dialog.html`
   - `.../src/fallback.html`
   - `.../assets/icon-16.png`
   - `.../assets/icon-32.png`
   - `.../assets/icon-80.png`
   - `.../assets/icon-64.png`
   - `.../assets/icon-128.png`

## Admin deployment
1. Go to Microsoft 365 Admin Center.
2. Open `Integrated apps`.
3. `Add-ins` -> `Upload custom apps` -> `Office Add-in`.
4. Upload `manifest.xml`.
5. Assign to `Entire organization`.
6. Save and wait for propagation.

## Notes
- Propagation can take time before all users see it.
- Users may need to restart Outlook to see the button quickly.
