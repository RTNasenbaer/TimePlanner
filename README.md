# TimePlanner - Football Training Planner

A professional training planning application for football coaches with advanced DOCX export capabilities.

## 🚀 Quick Start

### Requirements
- Windows 10/11
- Python 3.7 or higher
- Required packages (automatically installed)

### Installation
1. **Install Dependencies**: 
   ```
   pip install -r requirements.txt
   ```

2. **Run Application**:
   ```
   python main.py
   ```
   or double-click `run.ps1`

## ✨ Features

### Core Functionality
- **Interactive Timeline**: Drag and drop training segments
- **Right-Click Editing**: Edit segments directly on the timeline
- **Smart Tool Management**: Automatically combines equipment needs
- **Professional Export**: Generate DOCX training documents

### Advanced Features
- **Settings Management**: Configure trainer info, player count, etc.
- **Template-Based Export**: Uses your custom DOCX template
- **Smart Filename Generation**: `Trainingseinheit_Name_Theme_Date.docx`
- **Font Formatting**: Professional Calibri 14pt output
- **Excel Import/Export**: Compatible with spreadsheet workflows

### User Interface
- **Visual Timeline**: Clear segment representation with colors
- **Context Menus**: Right-click for quick actions
- **Scalable Interface**: Zoom in/out for better visibility
- **Professional Styling**: Modern, clean interface

## 📁 Essential Files

### Core Application
- `main.py` - Main application file
- `requirements.txt` - Python dependencies
- `settings.json` - User settings (auto-created)

### Template & Assets
- `Trainingseinheit_Name_StandardFile_Date.docx` - DOCX export template
- `icon.ico` - Application icon

### Build & Distribution
- `build-exe.bat` - Build executable (Windows batch)
- `build-exe.ps1` - Build executable (PowerShell)
- `create-distribution.ps1` - Create distribution package
- `run.ps1` - Run application script
- `setup.ps1` - Setup environment
- `TimePlanner.spec` - PyInstaller configuration

### Portable Version
- `TimePlanner-Portable/` - Standalone executable version

### Documentation
- `README.md` - This file
- `INSTALLATION.md` - Detailed installation guide
- `DOCX_Implementation_Summary.md` - Technical implementation details

## 🎯 How to Use

### 1. Basic Training Plan
1. Launch the application
2. Enter plan name and duration
3. Add segments using "Abschnitt hinzufügen"
4. Drag segments to reorder

### 2. Edit Segments
- **Right-click** on any segment
- Choose "Bearbeiten" to modify properties
- Choose "Löschen" to remove

### 3. Configure Settings
- Go to "Einstellungen" → "Einstellungen..."
- Set trainer name, player count, team info

### 4. Export Training Document
1. Go to "Datei" → "Als DOCX exportieren" (Ctrl+D)
2. Application uses the template automatically
3. Choose output filename
4. Professional document is generated

## 🛠️ Build Executable

To create a standalone executable:

```powershell
# Build executable
.\build-exe.ps1

# Create distribution package
.\create-distribution.ps1
```

## 📝 Technical Details

- **Framework**: PyQt5 for GUI
- **Export**: python-docx for Word document generation
- **Data**: JSON for settings, Excel for import/export
- **Graphics**: Custom painting for timeline visualization

## 🎉 Key Features Summary

✅ **Interactive Timeline** with drag & drop  
✅ **Right-click Editing** for quick modifications  
✅ **Professional DOCX Export** with template support  
✅ **Smart Tool Combination** (keeps highest quantities)  
✅ **Duration Display** instead of time slots  
✅ **Custom Filename Format** with trainer name and date  
✅ **Calibri 14pt Font** for professional appearance  
✅ **Settings Management** for trainer information  
✅ **Excel Compatibility** for data exchange  

Perfect for football coaches who want professional training documentation! ⚽