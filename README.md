# Archie Museum

A minimalist online art gallery for Archie's drawings.

## Deploy to GitHub Pages

1. Open PowerShell in this folder.
2. Run:

```powershell
.\deploy.ps1
```

3. Sign in to GitHub when prompted.
4. Open the printed URL on your phone after 1-3 minutes.

If the repository already exists, push manually:

```powershell
git remote add origin https://github.com/<your-username>/archie-museum.git
git push -u origin master
```

Then enable **GitHub Pages** in repository settings: branch `master`, folder `/ (root)`.
