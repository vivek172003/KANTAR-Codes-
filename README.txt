Install: py -m pip install -r requirements.txt
Run: py gdf_sanity_enhanced.py "C:\incoming" --filter-workbook "C:\UL Global DVP Scale up - Feedback Log.xlsx" -o "C:\GDF_Sanity_Report.xlsx"
The program reads CASEDATA, METADATA and EXISTENCE CSV files directly from folders or ZIPs.

BEAST_sheet_Names.py SharePoint setup:
1. Register a public-client application in Microsoft Entra ID and enable device-code/public-client sign-in.
2. Add Microsoft Graph delegated Files.Read.All permission and obtain any required administrator consent.
3. Set MS_TENANT_ID and MS_CLIENT_ID in the environment; do not put app secrets in this script.
4. Run: py BEAST_sheet_Names.py. Sign in when Microsoft displays the device-code prompt.

The script reads Excel files directly from the configured SharePoint folder link. The signed-in account must
already have access to that folder. If organizational policy blocks app registration or Graph consent,
use the synced OneDrive folder path instead.
