#!/bin/bash

echo "===================================="
echo "  TimePlanner - Automated Setup"
echo "===================================="
echo

# Check if conda is installed
if ! command -v conda &> /dev/null; then
    echo "ERROR: Conda is not installed or not in PATH."
    echo "Please install Miniconda or Anaconda first:"
    echo "https://docs.conda.io/en/latest/miniconda.html"
    echo
    exit 1
fi

echo "Checking conda installation..."
conda --version
echo

# Initialize conda for bash if needed
eval "$(conda shell.bash hook)" 2>/dev/null || {
    echo "Initializing conda for this shell session..."
    source "$(conda info --base)/etc/profile.d/conda.sh" 2>/dev/null || {
        echo "WARNING: Could not initialize conda. You may need to restart your terminal."
    }
}

# Remove existing environment if it exists
echo "Removing existing 'zeitplan' environment if it exists..."
conda env remove -n zeitplan -y 2>/dev/null
echo

# Create new conda environment
echo "Creating new conda environment 'zeitplan' with Python 3.11..."
conda create -n zeitplan python=3.11 -y
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to create conda environment."
    exit 1
fi
echo

# Activate environment and install packages
echo "Installing required packages..."
conda activate zeitplan
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to activate conda environment."
    exit 1
fi

# Install conda packages
echo "Installing conda packages..."
conda install pyqt pyqtgraph pandas openpyxl -y
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to install conda packages."
    exit 1
fi

# Install pip packages
echo "Installing pip packages..."
pip install PyQt5 qtawesome
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to install pip packages."
    exit 1
fi

echo
echo "===================================="
echo "  Setup completed successfully!"
echo "===================================="
echo
echo "To run the TimePlanner application:"
echo "1. Open terminal"
echo "2. Run: conda activate zeitplan"
echo "3. Run: python main.py"
echo
echo "Or use the provided run scripts:"
echo "- Windows: setup.bat"
echo "- Linux/macOS: ./setup.sh"
echo