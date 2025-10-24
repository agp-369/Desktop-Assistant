#!/bin/bash
# This script installs the dependencies for the desktop assistant.

# First, install the Python dependencies.
pip install -r requirements.txt

# For Linux users, PyAudio and pyautogui require some additional system packages.
# The following commands should work for Debian-based systems (like Ubuntu).
# If you are on a different Linux distribution, you may need to use a different package manager.
if [[ "$OSTYPE" == "linux-gnu"* ]]; then
    echo "Detected Linux. For PyAudio and pyautogui to work, you may need to install the following packages:"
    echo "sudo apt-get install portaudio19-dev python3-pyaudio scrot python3-tk python3-dev"
fi
