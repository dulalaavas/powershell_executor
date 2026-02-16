# 🚀 PowerShell Executor

A Python-based GUI application for executing PowerShell commands from a single, user-friendly interface. Streamline your Windows system administration tasks with pre-configured commands and custom scripts.

![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

## 📋 Table of Contents

- [Features](#-features)
- [Screenshots](#-screenshots)
- [Installation](#-installation)
- [Usage](#-usage)
- [Configuration](#-configuration)
- [Command Management](#-command-management)
- [Building Executable](#-building-executable)
- [Project Structure](#-project-structure)
- [Contributing](#-contributing)
- [License](#-license)

## ✨ Features

- 🎨 **Modern GUI Interface** - Clean, intuitive graphical user interface built with Python
- ⚡ **Quick Command Execution** - Execute PowerShell commands with a single click
- 📝 **Custom Commands** - Add, edit, and manage your own PowerShell commands
- 💾 **Command Library** - Pre-configured with common Windows administration tasks
- 🔧 **JSON Configuration** - Easy-to-edit command storage in JSON format
- 📊 **Real-time Output** - View command execution results instantly
- 🛡️ **Admin Detection** - Automatic detection of administrator privileges
- 🎯 **Categorized Commands** - Organize commands by categories for better management

## 📸 Screenshots

> Add screenshots of your application here

## 🔧 Installation

### Prerequisites

- Windows 7/8/10/11
- Python 3.7 or higher
- PowerShell 5.1 or higher

### Quick Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/dulalaavas/powershell_executor.git
   cd powershell_executor
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   python app_design.py
   ```

### Using Pre-built Executable

Download the latest release from the [Releases](https://github.com/dulalaavas/powershell_executor/releases) page and run the executable directly.

## 🎯 Usage

### Basic Usage

1. Launch the application
2. Select a command from the list or category dropdown
3. Click "Execute" to run the command
4. View the output in the results panel

### Running as Administrator

Some PowerShell commands require administrator privileges. To run the application with elevated rights:

1. Right-click on the executable or Python script
2. Select "Run as Administrator"
3. The application will display an admin indicator in the title bar

### Adding Custom Commands

You can add your own PowerShell commands by editing the `commands.json` file:

```json
{
  "commands": [
    {
      "name": "Check Disk Space",
      "category": "System Info",
      "command": "Get-PSDrive C | Select-Object Used,Free",
      "description": "Display disk space information for C: drive",
      "requires_admin": false
    },
    {
      "name": "Restart Service",
      "category": "Services",
      "command": "Restart-Service -Name 'ServiceName' -Force",
      "description": "Restart a Windows service",
      "requires_admin": true
    }
  ]
}
```

## ⚙️ Configuration

### commands.json

The `commands.json` file stores all available PowerShell commands:

```json
{
  "commands": [
    {
      "name": "Command Name",
      "category": "Category",
      "command": "PowerShell-Command",
      "description": "Brief description",
      "requires_admin": false
    }
  ]
}
```

**Field Descriptions:**
- `name`: Display name for the command
- `category`: Organization category
- `command`: The actual PowerShell command to execute
- `description`: Helpful description of what the command does
- `requires_admin`: Boolean flag indicating if admin privileges are needed

### config.json

Application settings and preferences:

```json
{
  "theme": "light",
  "auto_admin_detect": true,
  "log_output": true,
  "max_output_lines": 1000
}
```

## 📚 Command Management

### Pre-configured Command Categories

The application comes with pre-configured commands in the following categories:

- **System Information** - Get system details, hardware info, OS version
- **Network Tools** - Network diagnostics, IP configuration, connectivity tests
- **Service Management** - Start, stop, restart Windows services
- **Performance** - System optimization, cleanup tasks
- **Security** - Windows Defender, firewall settings
- **User Management** - User accounts, permissions
- **File Operations** - File search, bulk operations

### Creating Command Sets

You can create custom command sets for specific workflows:

1. Create a new JSON file (e.g., `my_commands.json`)
2. Use the same structure as `commands.json`
3. Load the custom set in the application settings

## 🏗️ Building Executable

To create a standalone executable using PyInstaller:

```bash
# Install PyInstaller
pip install pyinstaller

# Build the executable
pyinstaller --onefile --windowed --icon=icon.ico app_design.py

# The executable will be in the dist/ folder
```

### Advanced Build Options

```bash
# Include data files
pyinstaller --onefile --windowed \
  --add-data "commands.json;." \
  --add-data "config.json;." \
  --icon=icon.ico \
  app_design.py
```

## 📁 Project Structure

```
powershell_executor/
│
├── app_design.py          # Main application file
├── commands.json          # Command definitions
├── config.json           # Application configuration
├── requirements.txt      # Python dependencies
│
├── dist/                 # Built executables (generated)
│   └── app_design.exe
│
├── assets/              # Images, icons (optional)
│   └── icon.ico
│
└── README.md           # This file
```

## 🛠️ Development

### Setting Up Development Environment

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run in development mode
python app_design.py
```

### Adding New Features

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

- 🐛 Report bugs and issues
- 💡 Suggest new features
- 📝 Improve documentation
- 🔧 Submit pull requests

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on our code of conduct and the process for submitting pull requests.

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Inspired by tools like Chris Titus Tech's Windows Utility
- Built with Python and PowerShell
- Special thanks to the open-source community

## 📧 Contact

Dulal Aavas - [@dulalaavas](https://github.com/dulalaavas)

Project Link: [https://github.com/dulalaavas/powershell_executor](https://github.com/dulalaavas/powershell_executor)

## 🔗 Related Projects

- [Chris Titus Tech Windows Utility](https://github.com/ChrisTitusTech/winutil)
- [PowerShell Gallery](https://www.powershellgallery.com/)
- [Windows Terminal](https://github.com/microsoft/terminal)

## ⚠️ Disclaimer

This tool executes PowerShell commands on your system. Always review commands before executing them, especially when running with administrator privileges. The authors are not responsible for any damage caused by misuse of this tool.

---

**Made with ❤️ by [Dulal Aavas](https://github.com/dulalaavas)**
