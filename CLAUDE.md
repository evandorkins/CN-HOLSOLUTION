# CN-HOLSOLUTION Project Instructions

## UKG SDM Import Packaging

When preparing configuration files for UKG import, follow this packaging process:

### Package Structure
```
sdm-MM-DD-YYYY/
├── ExportConfig.json
└── {ConfigType}/
    └── response.json
```

### Steps

1. **Create dated folder** in project root:
   - Format: `sdm-MM-DD-YYYY` (e.g., `sdm-12-14-2024`)

2. **Copy ExportConfig.json** from `fullHOLSolution/ExportConfig.json` to the dated folder

3. **Create config subfolder** named after the original response.json parent folder:
   - Example: If source is `fullHOLSolution/WSAPayCode/response.json`, create subfolder `WSAPayCode/`
   - Common config types: `WSAPayCode`, `WSAHolidayCreditRule`, `WSAAccrualPolicy`, `WSAAccrualProfile`, etc.

4. **Copy response.json** (modified or new) into the config subfolder

5. **Create zip file** of the entire sdm folder:
   ```bash
   zip -r sdm-MM-DD-YYYY.zip sdm-MM-DD-YYYY/
   ```

### Config Types Reference
| Folder Name | Description |
|-------------|-------------|
| WSAPayCode | Pay Code definitions |
| WSAHolidayCreditRule | Holiday Credit Rules |
| WSAAccrualPolicy | Accrual Policies |
| WSAAccrualProfile | Accrual Profiles |
| WSACfgAccrualCode | Accrual Codes |
| WSACustomDate | Custom Dates |
| WSADatePattern | Date Patterns |
| WSALimit | Limits |
| WSABalanceCascade | Balance Cascades |
| WSABalanceCascadeGroup | Balance Cascade Groups |
| WSAContributingPayCodeRule | Contributing Pay Code Rules |
| WSAContributingShiftRule | Contributing Shift Rules |
| WSAHoliday | Holidays |
| APIHolidayProfile | Holiday Profiles |
| EmploymentTerm | Employment Terms |

### Example Command
```bash
# Package WSAPayCode for import
mkdir -p sdm-12-14-2024/WSAPayCode
cp fullHOLSolution/ExportConfig.json sdm-12-14-2024/
cp fullHOLSolution/newpaycodes/response.json sdm-12-14-2024/WSAPayCode/
zip -r sdm-12-14-2024.zip sdm-12-14-2024/
```

## Project Structure
- `fullHOLSolution/` - Source configuration exports from UKG
- `sdm-*/` - Packaged imports ready for UKG
- Response.json files contain configuration items in UKG export format
