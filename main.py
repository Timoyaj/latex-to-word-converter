#!/usr/bin/env python3
"""LaTeX to Word Converter - Main Entry Point"""
import sys
import os

# Ensure project root is on path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui.app import main

if __name__ == "__main__":
    main()
