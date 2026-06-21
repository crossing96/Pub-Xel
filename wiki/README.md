# Pub-Xel Wiki (source)

User guide pages live here in the main repo. The **GitHub Wiki tab** uses a separate git repo — sync is automatic or one command away.

## Automatic (recommended)

On push to `main`, the [Sync wiki](../../.github/workflows/sync-wiki.yml) workflow copies `wiki/*.md` to the GitHub Wiki repo.

**One-time setup:** open the [Wiki tab](https://github.com/crossing96/Pub-Xel/wiki) on GitHub and create any page (e.g. **Home**) once so the `.wiki` repository exists. After that, pushes to `wiki/` update the Wiki tab.

You can also run the workflow manually: **Actions → Sync wiki → Run workflow**.

## Manual sync (local)

```powershell
.\scripts\sync_wiki.ps1
```

Copy only, no push:

```powershell
.\scripts\sync_wiki.ps1 -Push:$false
```

Or with Python only:

```powershell
python scripts/sync_wiki.py wiki path\to\Pub-Xel.wiki
```

## Notes

- Edit files in `wiki/` here, commit to the main repo, then push — the workflow handles the rest.
- Internal links use `.md` in source (e.g. `[Installation](Installation.md)`); sync strips `.md` for GitHub Wiki.
- `README.md` in this folder is for maintainers only and is not published to the Wiki tab.
