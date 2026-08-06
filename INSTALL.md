# Installing LudicrousSpeed

LudicrousSpeed ships as a prebuilt Windows Excel add-in. No build tools required.

1. Grab the latest build from the **[Releases page](https://github.com/Grant-Schwartz/WarpSpeed/releases/latest)** — download the `ludicrous-windows-*.zip` under Assets.
2. Optional but recommended: right-click the downloaded `.zip` → **Properties** → check **Unblock** → **OK**. This skips step 4 below entirely.
3. Unzip it anywhere.
4. Double-click **Install.cmd**. If Windows says "Windows protected your PC," click **More info > Run anyway** — this build isn't code-signed yet.
5. Open Excel. LudicrousSpeed should already be on the ribbon — the installer registers it to auto-load, not just trusts the folder.
   If it isn't showing: **File > Options > Add-ins > Manage: Excel Add-ins > Go... > Browse...**, then select `LudicrousSpeed.xll` from the path `Install.cmd` printed.

Everything the installer does is per-user (no admin rights, no UAC prompt).

## Uninstalling

From the same unzipped folder, double-click **Uninstall.cmd**.

This removes the copied files, the Trust Center entry, and the auto-load registration, and reminds you to also remove the add-in from Excel's Add-ins list if it's still showing there.

## Heads up

This is a beta build — it isn't code-signed, and a fresh release is cut automatically from `main` on every push, so expect rough edges. See the [README](README.md) for how the engine actually works.
