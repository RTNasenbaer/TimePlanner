#!/bin/bash
echo "Starting TimePlanner..."

# Initialize conda for bash if needed
eval "$(conda shell.bash hook)" 2>/dev/null || {
    source "$(conda info --base)/etc/profile.d/conda.sh" 2>/dev/null || {
        echo "ERROR: Could not initialize conda."
        exit 1
    }
}

conda activate zeitplan
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to activate conda environment 'zeitplan'."
    echo "Please run ./setup.sh first to create the environment."
    exit 1
fi

python main.py