# 📦 Auto-Unzipper

![License](https://img.shields.io/github/license/Nrentzilas/Auto-Unzipper)
![GitHub issues](https://img.shields.io/github/issues/Nrentzilas/Auto-Unzipper)

A simple, user-friendly desktop application that automatically extracts compressed archives from your downloads folder (or any folder you choose).



## ✨ Features

- 🔍 Monitors a folder for new archive files (`.zip`, `.rar`, `.7z`)
- 🚀 Automatically extracts archives when detected
- 📂 Custom extraction destination
- 🗑️ Option to delete original archives after successful extraction
- ⏱️ Configurable monitoring interval
- 📝 Activity logging
- 🛠️ Customizable settings that persist between sessions

## 🛠️ Requirements

- Python 3.6+
- PyQt6
- 7-Zip (installed on your system)

## ⚙️ Installation

### Method 1: From Source

1. Clone the repository:
```bash
git clone https://github.com/Nrentzilas/Auto-Unzipper.git
cd Auto-Unzipper
```

2. Install the required dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python auto_unzipper.py
```

### Method 2: Executable (Windows)

You can download the pre-built executable from the [Releases](https://github.com/Nrentzilas/Auto-Unzipper/releases) page.

## 📦 Building the Executable

To build the executable yourself using PyInstaller:

1. Install PyInstaller:
```bash
pip install pyinstaller
```

2. Create the executable:
```bash
pyinstaller --name "Auto-Unzipper" --windowed --icon=icon.ico --onefile auto_unzipper.py
```

The executable will be created in the `dist` directory.

## 📝 Usage

1. Start the application
2. Configure the settings:
   - **Monitor folder**: The folder to watch for new archives (default: Downloads folder)
   - **Extract to**: Where extracted files should be placed
   - **7-Zip executable**: Path to 7z.exe
   - **Check interval**: How often to scan for new files (in seconds)
   - **Delete archives after extraction**: Option to remove archives after successful extraction
   - **Auto-start monitoring on launch**: Start monitoring automatically when the app opens
3. Click "Save Settings" to store your preferences
4. Click "Start Monitoring" to begin the automatic extraction process
5. The Activity Log will show all operations and any errors

## 🤝 Contributing

Contributions are welcome! Here's how you can contribute:

1. Fork the repository
2. Make your changes
3. Open a Pull Request

Please make sure to update tests as appropriate and follow the code style.

## 🐛 Issues and Bug Reports

If you encounter any problems or have suggestions, please [open an issue](https://github.com/Nrentzilas/Auto-Unzipper/issues) on GitHub.

When reporting issues, please include:
- A clear description of the problem
- Steps to reproduce the issue
- Expected behavior
- Screenshots if applicable
- Your operating system and Python version

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgements

- [7-Zip](https://www.7-zip.org/) for the extraction functionality
- [PyQt6](https://www.riverbankcomputing.com/software/pyqt/) for the GUI framework