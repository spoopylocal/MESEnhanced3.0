# MES Enhanced 3.0

A comprehensive Chrome/Firefox extension that enhances the MES (Manufacturing Execution System) interface with powerful productivity tools.

## Features

### Quick Copy Buttons
- **Part Numbers (PN)**: Individual copy buttons next to each PN
- **Copy All PNs**: Bulk copy all part numbers from a work order
- **Copy All PNs + QTY**: Copy part numbers with quantities
- **LPN Copy**: Copy License Plate Numbers
- **LPN + Location**: Copy LPNs with their locations
- **Work Order (WO)**: Quick copy work order number
- **Service Template (SR)**: Copy service template name

### MO Tracking Tools (Tools Menu)

#### 1. MO PRM/INFO
- Enter multiple MO (Move Order) numbers (comma, space, or newline separated)
- Automatically fetches and displays:
  - Work Order Number
  - Build Name (from PRM)
  - Build Due Date
- Auto-assigns PRM if missing (finds latest published PRM from MSF code)
- Copies formatted output to clipboard
- Shows progress with toast notifications

**Example Input:**
```
MO156913878, MO156913879
```

**Example Output:**
```
MO156913878    WO123456    Dell PowerEdge R740        15-Mar
MO156913879    WO123457    HPE ProLiant DL380         20-Mar
```

#### 2. MO STATUS
- Paste a list of MO numbers
- Checks current status for each:
  - Open (not in system yet)
  - Picking (being picked)
  - Inventoried (in inventory)
  - DRP location status
- Identifies problematic orders (in DRP but being worked on)
- Copies summary to clipboard
- Detailed console output with warnings

### Serial Checkers (Tools Menu)

#### 3. LSC Serial Checker (LinkedIn Serial Checker)
- Scans work order for serial number completeness
- Checks for required fields: Asset Number, Rack Position, Warranty Start, Warranty End
- Classifies devices as Main Device or Node
- Shows missing attributes in organized modal
- Copy-friendly output format
- Uses `CFG_LTO` document type

#### 4. IBC Serial Checker (Inbound Checker)
- Enhanced serial checker for inbound validation
- Required base fields: Asset Number, Rack Position
- Conditional fields: MSF Part, MSF Asset (when applicable)
- Uses `WIP_JOB` document type
- Ignores MAC address fields from validation

### Asset Management

#### ASSET Button
- Downloads asset sheet attachment automatically
- Searches for attachments with keywords: "asset qc", "qc", "asset sheet", "asset"
- Prioritizes QC sheets
- Shows loading indicator during download
- Provides feedback if no asset sheet found

### Placard Features
- Auto-filled placard downloads
- Pre-populates fields from work order:
  - MES Work Order
  - Move Orders
  - Order Type
  - GICLAB Ticket #
  - Power Type (fetched from PRM)
  - Customer PO
  - And more...
- Intercepts "Generate Placard" menu option to use enhanced version

## Installation

1. Open Chrome/Firefox
2. Navigate to Extensions page:
   - Chrome: `chrome://extensions`
   - Firefox: `about:addons`
3. Enable "Developer mode"
4. Click "Load unpacked"
5. Select the `MESEnhanced3.0` folder

## Usage

### On MES Work Order Pages:
1. **Copy buttons** appear next to Part Numbers and other data
2. **Tools menu** (gear icon) in header provides access to:
   - MO PRM/INFO
   - MO STATUS  
   - LSC Serial Checker
   - IBC Serial Checker
3. **ASSET button** in work order header for quick asset sheet download
4. **SR button** to copy service template name
5. **WO button** to copy work order number

### Tools Menu:
Click the gear icon in the top-right header to access all tools.

## Technical Details

### API Endpoints Used:
- `https://apirouter.apps.wwt.com/api/forward/mes-api/*`
- `https://apirouter.apps.wwt.com/prm-api/*`
- `https://apirouter.apps.wwt.com/api/forward/serial-attributes/*`
- `https://apirouter.apps.wwt.com/api/forward/attachments/*`

### Permissions:
- `clipboardWrite`: Copy data to clipboard
- `downloads`: Download asset sheets and placards
- `webRequest`: Enhanced API access for serial scanning

### Browser Compatibility:
- Chrome/Chromium
- Firefox (with WebExtensions)

## Changelog

### Version 3.0.0 (Current)
- Integrated LSC and IBC serial checkers
- Fixed asset sheet download with better error handling
- Updated MO PRM/INFO with improved error checking and batch processing
- Added comprehensive toast notifications
- Improved UI consistency across all features
- Better clipboard handling with fallbacks
- Enhanced modal styling for serial scan results

### Version 2.0
- Added placard auto-fill
- Service template copy
- MO tracker functionality

### Version 1.0
- Basic PN copy buttons
- LPN copy functionality
- Asset download

## Support

For issues or feature requests, contact the development team or create an issue in the project repository.

## License

Internal tool for WWT use.
