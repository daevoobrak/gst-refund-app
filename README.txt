GST REFUND WORKING GENERATOR
=============================

REQUIREMENTS
------------
- Python 3.8 or later  (https://www.python.org/downloads/)
  Windows: during install, TICK "Add Python to PATH"
- Internet connection for first run (downloads ~15 MB of packages)
- No other software needed

HOW TO RUN
----------

Windows:
  Double-click  run.bat
  The browser will open automatically at http://localhost:5050

Mac / Linux:
  Open Terminal in this folder and run:
    chmod +x run.sh && ./run.sh

HOW TO USE
----------
1. Upload GSTR-2B Excel  (file name starts with the period, e.g. 012025_…xlsx)
2. Upload GSTR-1 PDF     (downloaded from GST portal)
3. Upload GSTR-3B PDF    (downloaded from GST portal)
4. Click "Generate Refund Working"
5. The output Excel is downloaded automatically.

OUTPUT FILE
-----------
The generated Excel contains 4 sheets matching the Refund Working format:
  • Post Audit Sheet       – identity, refund claimed, Rule 89(4) max refund
  • MAIN CALCULATION SHEET – consolidated ITC comparison (GSTR-3B vs GSTR-2B)
  • Turnover Sheet         – period-wise GSTR-3B & GSTR-1 turnover breakdown
  • ITC Sheet              – 4(A)(1-5), 4(B), 4(C) from GSTR-3B and GSTR-2B

STOPPING THE SERVER
-------------------
Press Ctrl+C in the terminal / command prompt window.
