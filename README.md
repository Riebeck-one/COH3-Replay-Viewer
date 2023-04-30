# COH3 Replay Viewer

COH3 Replay Viewer is a Python-based application that allows you to view and manage your Company of Heroes 3 replay files. It provides an easy-to-use interface to track, save, browse, launch, rename, and delete replays. The program also supports the 'coh3-replay-enhancements' mod by Janne252.

## Features

- Automatically saves replay files at the end of a game of Company of Heroes 3
- Browse and view your Company of Heroes 3 replay files in a user-friendly interface
- Launch replays directly from the application
- Rename replay files for easier organization
- Delete unwanted replays
- Copy the 'dofile('replay-enhancements/init.scar')' command to the clipboard when launching a replay (useful when using the 'coh3-replay-enhancements' mod)
- Automatically update the replay list when new replays are added or existing ones are modified

## Installation

1. Download the latest release from the [GitHub Releases](https://github.com/Maxinova/COH3-Replay-Viewer/releases) page.
2. Extract the contents of the zip file to a folder of your choice.
3. Ensure that the 'Assets' folder is located in the same directory as the COH3ReplayViewer executable.

## Requirements

To run the COH3 Replay Viewer, you will need the following Python libraries:

- `time`
- `datetime`
- `shutil`
- `os`
- `sys`
- `winreg`
- `tkinter`
- `threading`
- `tkscrolledframe`
- `ctypes.wintypes`
- `PIL (Pillow)`
- `pyperclip`
- `webbrowser`
- `psutil`
- `subprocess`

If you are running the Python script directly, make sure to install the required libraries using pip:

```
pip install tkscrolledframe pillow pyperclip psutil
```

## Usage

1. Run the COH3ReplayViewer executable or the `COH3ReplayViewer.py` script.
2. The application will display a list of your Company of Heroes 3 replays, along with information such as map, duration, and players.
3. Use the buttons to launch, rename, or delete replays as needed.
4. If you have the 'coh3-replay-enhancements' mod installed, you can enable the "Mode 'COH3-Replay-Enhancement'" option to automatically copy the required command to the clipboard when launching a replay.

## Contributions

Special thanks to the following contributors:

- squareRoot17
- ewerybody

The 'coh3-replay-enhancements' mod used in this application was created by Janne252. You can find more information and download the mod from the [GitHub repository](https://github.com/Janne252/coh3-replay-enhancements).

## License

This project is licensed under the GNU GPLv3 License. See the [LICENSE](LICENSE.md) file for details.
