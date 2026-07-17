# Building the Windows .exe

This folder is a ready-to-use build project for turning `dicom_receiver_pusher.py`
into a Windows executable, `RAppsDicomReceiverPusher.exe`.

Because PyInstaller produces platform-specific binaries, the `.exe` must be
built **on Windows** (or on a Windows CI runner) — it can't be built on Linux/Mac.
The included GitHub Actions workflow handles this for you automatically.

## Option A — Build automatically with GitHub Actions (no Windows PC needed)

1. Create a new GitHub repository (public or private).
2. Upload the entire contents of this folder to the repo, preserving the
   `.github/workflows/build-exe.yml` path.
3. Push to the `main` branch, or open the repo's **Actions** tab and manually
   run the "Build Windows EXE" workflow (`workflow_dispatch`).
4. When the run finishes, open it and download the `RAppsDicomReceiverPusher-windows`
   artifact — it contains your `.exe`.

## Option B — Build it yourself on a Windows machine

1. Install Python 3.11+ from python.org.
2. Open a command prompt in this folder and run:
   ```
   pip install -r requirements.txt
   pyinstaller dicom_app.spec
   ```
3. The finished executable will be at `dist\RAppsDicomReceiverPusher.exe`.

## Notes

- `logo.png` and `lucide_icons_data.py` are bundled automatically — keep them
  next to `dicom_receiver_pusher.py` when building.
- The app generates `sopclass.ini`, log files, and encrypted config files
  (`rec.enc` / `push.enc`) next to wherever the `.exe` is run from, so keep
  the `.exe` in its own folder rather than a shared/system directory.
- If Windows SmartScreen flags the unsigned `.exe` on first run, choose
  "More info" → "Run anyway" — this is expected for unsigned executables and
  isn't specific to this app.
- If you later want a custom taskbar icon, add an `.ico` file to this folder
  and set `icon="yourfile.ico"` in `dicom_app.spec`.
