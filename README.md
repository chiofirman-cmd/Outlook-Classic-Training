# Outlook Classic PST Archive Training

This repository builds a five-minute, English-language, 1920x1080 MP4 training simulation for Microsoft Outlook Classic. It uses a scripted Outlook-style interface rather than controlling or recording a real Outlook installation.

## Included training topics

1. Create a monthly PST file using `name_MMYY.pst` (example: `Archive_1126.pst`).
2. Set the PST as the default data file.
3. Configure a general incoming-message rule (not sender-specific).
4. Configure a general sent-message rule (not recipient-specific).
5. Test and verify Inbox and Sent Items archive destinations.

The simulated PST contains `Inbox` and `Sent Items` folders. Rule names are functional labels (`Inbox Archive Rule` and `Sent Items Archive Rule`), not month-based PST filenames.

## Build with GitHub Actions

1. Create a new GitHub repository.
2. Upload the contents of this ZIP to the repository root (the `.github` folder must remain at the root).
3. Commit the files to the `main` branch.
4. Open the repository's **Actions** tab.
5. Select **Build Outlook Classic Training Video**.
6. Click **Run workflow** (or let the workflow run after pushing to `main`).
7. Wait for the green check mark.
8. Open the completed workflow run and download **Outlook-Classic-Training-MP4** under **Artifacts**.

## Local build requirements

- Python 3.11+
- Pillow
- FFmpeg
- espeak-ng

Run `python -m pip install pillow` and then `python build.py` after installing FFmpeg and espeak-ng. The MP4 is written to `output/Outlook_Classic_Training.mp4`.

## Important Outlook behavior note

This is a visual training simulation, not a live Outlook configuration utility. Actual Outlook rule capabilities depend on account type and Outlook version. Moving messages to a local PST is commonly a client-side action and may require classic Outlook to be running. Setting a PST as the default data file does not move a Microsoft 365 or Exchange mailbox or change server-side mail delivery. Verify the exact rule action in the target Outlook environment before using it operationally.
