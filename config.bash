# Ledger/config.sh
# Configuration for the Ledger Executable Project (Function-bound Sheet)

# Safely target STAGING or PRODUCTION:
# At deploy time, CLASP relies on deployment id to target the correct bound project.
# At runtime, MenuSetup relies on the file-relative SpreadsheetApp.getActiveSpreadsheet().
# These IDs are stable, permanent endpoints (AKfycb...) obtained once
# via 'clasp deployments' and the Apps Script GUI.

# PRODUCTION DEPLOYMENT ID: The ID linked to your live production spreadsheet.
export DEP_PROJECT_DEPLOYMENT_ID_PROD="AKfycbyg2iW4qtcZyn4xyrPgFYhdzYvCunDCs_Lkebcy3XRC4fI8CtCBMnveqC1I-9e088Q"

# STAGING DEPLOYMENT ID: The ID linked to your test/staging spreadsheet copy.
export DEP_PROJECT_DEPLOYMENT_ID_STAGING="AKfycbzIVCk-tWYV5-n6LJ9pYiUu6O5MaduonRN1pBw1IUmtbjxpV1kjkAy3s1hv3nHq4KEq"

# 3. TRUST_UTILS_DEP_SYMBOL: The User Symbol defined in Ledger's appsscript.json.
export TRUST_UTILS_DEP_SYMBOL="TrustUtilsLib"
