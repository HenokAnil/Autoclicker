# Autoclicker 

(This project was just created for fun cause i was bored one night)

This project provides a simple autoclicker script written in Python. It allows you to automate mouse and keyboard clicks with customizable intervals and start/stop controls.

## Installation

### Requirements
- Python 3.x
- pynput library

Install dependencies:
```bash
pip install pynput
```

## Usage

### Running from Source
1. Install the required dependencies (see above).
2. Run the application:
   ```bash
   python autoclicker_app.py
   ```
3. Follow the on-screen instructions to configure click intervals and control keys.

### Using the Executable (.exe)
If you prefer not to install Python, you can use the pre-built executable:

1. Download `Autoclicker.exe` from the `dist` folder
2. Double-click to run - no installation needed!
3. The application will launch with the full GUI

## Building Your Own Executable

To build the .exe file yourself:

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```

2. Build the executable:
   ```bash
   pyinstaller --onefile --windowed --name "Autoclicker" autoclicker_app.py
   ```

3. The executable will be created in the `dist` folder

