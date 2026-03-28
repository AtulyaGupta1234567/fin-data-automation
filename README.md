# Live Mutual Fund NAV Tracker (VBA)
A high-speed Excel automation tool that synchronizes live Mutual Fund data directly into your spreadsheet. By simply entering Scheme Codes in Column A, the engine fetches the corresponding Scheme Name, Live NAV, and NAV Date instantly.

 # Performance Impact
Before: Manual tracking and data entry used to take 4 to 5 hours of intensive work per update.

After: The entire process is now fully automated and completes in under 2 seconds, regardless of the number of schemes.

 # Key Features
High-Velocity Retrieval: Processes large portfolios in seconds using optimized XMLHTTP requests.

Data Accuracy: Eliminates human error by pulling directly from official financial data sources.

Dynamic Scaling: Automatically detects your list length in Column A and updates every row.

Zero UI Lag: Engineered to fetch data efficiently without freezing Excel.

 # How to Use
Open your Excel workbook and press Alt + F11 to open the VBA Editor.

Go to Insert > Module.

Copy the code from mf_nav_updater.vba in this repository and paste it into the module.

Go to Tools > References and ensure Microsoft XML, v6.0 is checked.

Enter your Scheme Codes in Column A and run the macro.
# Quick Start
1. Download the `sample_portfolio_template.xlsx` from this repository.
2. Open the file and follow the **How to Use** steps above to add the VBA code.
3. Run the macro to see the `Scheme Name`, `NAV`, and `Date` columns update in under 2 seconds.
