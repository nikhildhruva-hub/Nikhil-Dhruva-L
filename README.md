# Google Apps Script in VS Code

This folder is ready to be linked to a Google Apps Script project with `clasp`.

## Link an existing Apps Script project

1. Open the VS Code terminal in this folder.
2. Sign in once:

```powershell
clasp login
```

3. In Google Apps Script, open **Project Settings** and copy the **Script ID**.
4. Link this folder:

```powershell
clasp clone YOUR_SCRIPT_ID
```

Replace `YOUR_SCRIPT_ID` with the copied value. This creates `.clasp.json` and downloads the Apps Script files here.

## Daily workflow

```powershell
clasp pull   # download the latest Apps Script files
clasp push   # upload local VS Code changes
clasp open   # open the project in the Apps Script editor
```

For an Apps Script web app, deploy after pushing:

```powershell
clasp deploy
```

Do not commit `.clasp.json` if it contains a project ID you do not want to share. The `index.html` file in this folder is currently empty; `clasp clone` will add the actual Apps Script files for the linked project.
