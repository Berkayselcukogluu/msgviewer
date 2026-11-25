# MSG Viewer

MSG Viewer is a lightweight, 100% offline desktop application for opening, reading, and converting Microsoft Outlook `.msg` files locally. It processes private documents directly on your computer without requiring an internet connection or uploading data to external services.

## Clone the Repository

```bash
git clone https://github.com/Berkayselcukogluu/msg-viewer.git
cd msg-viewer
```

## Installation

### Option 1: Standalone Executable (Windows)
No installation or Python setup is required. Launch `MSGViewer.exe` directly.

### Option 2: Run from Source
Requires Python 3.10 or newer:

```bash
pip install -r requirements.txt
python msg_viewer.py
```

## Usage and Features

- **Open Files and Folders:** Add individual `.msg` files via `+ Files`, scan directories recursively via `Folder`, pass file paths from the command line, or drag and drop files directly into the left panel.
- **Readable Dark Interface:** Automatically converts dark or black inline `.msg` font colors and light table backgrounds into clean white text for comfortable reading on the dark theme.
- **Unified Content View:** Displays general message details (`Subject`, `From`, `To`, `CC`, `BCC`, `Date`, `File`, and `Attachments`) directly above the email body inside the **Content** tab.
- **Persistent Local Cache:** Automatically stores opened `.msg` files in a local cache (`%LOCALAPPDATA%\MSGViewer\cache`). Previously opened emails and their attachments remain available even if the original files are deleted from your computer, and can be restored anytime using `Save .MSG`.
- **Search and Filter:** Instantly filter loaded messages by subject, sender, recipient, filename, or plain-text content.
- **Export and Attachments:** Save individual or all attachments from the **Attachments** tab, or convert single and batch `.msg` files into standard `.eml` format.
