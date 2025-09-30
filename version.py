"""
Version information for TimePlanner
This file is automatically updated by the build system.
"""

__version__ = "1.0.0"
VERSION = "1.0.0"
BUILD_DATE = ""
GIT_COMMIT = ""

def get_version():
    """Get the current version string"""
    return __version__

def get_version_info():
    """Get detailed version information"""
    return {
        'version': __version__,
        'build_date': BUILD_DATE,
        'git_commit': GIT_COMMIT
    }

def get_version_string():
    """Get a formatted version string for display"""
    if BUILD_DATE:
        return f"TimePlanner v{__version__} (Built: {BUILD_DATE})"
    else:
        return f"TimePlanner v{__version__}"