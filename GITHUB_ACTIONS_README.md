# TimePlanner GitHub Actions Setup

This repository now has automated builds and releases set up using GitHub Actions. Here's how it works:

## 🚀 Automatic Releases

### When releases are created:
- **Every commit to the `main` branch** automatically triggers a new release
- **Pull requests** are tested but don't create releases

### Release versioning:
- Uses date-based versioning: `YYYY.MM.DD`
- If multiple releases happen on the same day: `YYYY.MM.DD.N` (where N is the commit count for that day)
- Examples: `2024.09.30`, `2024.09.30.2`, `2024.09.30.3`

## 📦 What gets built and released:

Each release includes:

1. **`TimePlanner-vX.X.X.exe`** - Standalone executable
   - Single file, no installation required
   - Includes all dependencies
   - Can be run directly on Windows 10+

2. **`TimePlanner-Portable-vX.X.X.zip`** - Portable package
   - Contains the executable renamed to `TimePlanner.exe`
   - Includes language files and other resources
   - Includes `VERSION.txt` with build information
   - Perfect for distribution or running from USB drives

## 🔧 Build Process

The automated build process:

1. **Sets up Python 3.11** in a Windows environment
2. **Installs all dependencies**: PyQt5, pandas, openpyxl, python-docx, pyinstaller, pillow
3. **Creates version information** with build date and commit hash
4. **Generates application icon** automatically
5. **Builds executable** using PyInstaller with optimized settings
6. **Creates portable distribution** with all necessary files
7. **Uploads both formats** as release assets

## 📋 Workflow Files

### `.github/workflows/release.yml`
- **Triggers**: Push to `main` branch
- **Creates**: GitHub releases with executable and portable ZIP
- **Includes**: Automatic changelog generation from commit messages

### `.github/workflows/test-build.yml`
- **Triggers**: Pull requests to `main` branch, manual dispatch
- **Purpose**: Test that the build works without creating a release
- **Output**: Test executable artifact (kept for 7 days)

## 🛠️ Local Development

### Updated Build Scripts

**`build-ci.ps1`** - Enhanced build script that works both locally and in CI:
```powershell
# Basic build
.\build-ci.ps1

# Build with specific version
.\build-ci.ps1 -Version "1.2.3"

# Test build (doesn't create portable distribution)
.\build-ci.ps1 -Test

# CI mode (no interactive prompts)
.\build-ci.ps1 -CI -Version "1.2.3"
```

### Version Management

The `version.py` file is automatically created/updated during builds with:
- Version number
- Build date and time
- Git commit hash (in CI builds)

The application window title will show: `TimePlanner v1.2.3 (Built: 2024-09-30 14:30:00) - Project Name`

## 🔄 How to Release

### Automatic (Recommended):
1. Make your changes
2. Commit and push to `main` branch
3. GitHub Actions automatically builds and creates a release
4. Release appears in the "Releases" section with:
   - Generated changelog from commit messages
   - Both executable formats
   - Installation instructions

### Manual:
- You can still use the existing `build-exe.ps1` script for local builds
- Use `build-ci.ps1` for more control over the build process

## 📝 Release Notes

Release notes are automatically generated from commit messages between releases. To get better release notes:

- Write clear, descriptive commit messages
- Use conventional commit format if desired:
  - `feat: add new feature`
  - `fix: resolve issue with...`
  - `docs: update documentation`
  - `refactor: improve code structure`

## 🔍 Monitoring Builds

- Check the **"Actions"** tab in GitHub to see build status
- Failed builds will show error details
- Successful builds automatically create releases
- Test builds (from PRs) create artifacts you can download

## 🚨 Troubleshooting

### Build fails:
1. Check the Actions tab for error details
2. Common issues:
   - Missing dependencies (should auto-install)
   - Python syntax errors
   - Missing files (icon, language files, etc.)

### Release not created:
- Only pushes to `main` branch create releases
- Pull requests only create test builds
- Check that the workflow has write permissions

## 🎯 Benefits

- **Zero-maintenance releases**: Just commit to main
- **Consistent builds**: Same environment every time
- **Professional releases**: With changelogs and proper versioning
- **Multiple formats**: Both standalone and portable versions
- **Quality assurance**: PRs are automatically tested

## 📚 Files Added/Modified

### New files:
- `.github/workflows/release.yml` - Main release workflow
- `.github/workflows/test-build.yml` - PR testing workflow  
- `build-ci.ps1` - Enhanced build script
- `version.py` - Version management system

### Modified files:
- `main.py` - Now shows version in window title
- Updated to use version information from `version.py`

---

**Note**: The first release after setting this up will include all commit history. Subsequent releases will only show changes since the last release.