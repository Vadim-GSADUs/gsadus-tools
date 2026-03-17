"""Double-click launcher — Windows runs .pyw files silently (no console window)."""
import os
import sys

# Ensure imports resolve correctly regardless of where this file is launched from
_here = os.path.dirname(os.path.abspath(__file__))
os.chdir(_here)
sys.path.insert(0, _here)

from app import main
main()
