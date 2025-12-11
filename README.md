# CN-HOLSOLUTION

UKG ProWFM Holiday Credit Solution Configuration Documentation and Analysis

## Overview

This repository contains the configuration exports and documentation for the Holiday Credit Solution in UKG ProWFM. The solution manages holiday credit accruals, eligibility rules, and forfeit processing for both Exempt and Non-Exempt employees.

## Repository Structure

```
├── fullHOLSolution/           # Main JSON configuration exports from UKG ProWFM
│   ├── APIHolidayProfile/     # Holiday Profiles (7 profiles)
│   ├── EmploymentTerm/        # Employment Terms (9 terms)
│   ├── WSAAccrualPolicy/      # Accrual Policies (223 policies)
│   ├── WSAAccrualProfile/     # Accrual Profiles (110 profiles)
│   ├── WSABalanceCascade/     # Balance Cascades (18 cascades)
│   ├── WSABalanceCascadeGroup/# Balance Cascade Groups (2 groups)
│   ├── WSACfgAccrualCode/     # Accrual Codes/Banks (9 codes)
│   ├── WSAContributingPayCodeRule/ # Contributing Pay Code Rules (18 rules)
│   ├── WSAContributingShiftRule/   # Contributing Shift Rules (9 rules)
│   ├── WSACustomDate/         # Custom Dates (9 dates)
│   ├── WSADatePattern/        # Date Patterns (36 patterns)
│   ├── WSAHoliday/            # Holidays (14 holidays)
│   ├── WSAHolidayCreditRule/  # Holiday Credit Rules (36 rules)
│   ├── WSALimit/              # Limits (9 limits)
│   └── WSAPayCode/            # Pay Codes (55 codes)
├── generate_documentation.py  # Python script to generate Excel documentation
├── HolidayCreditSolution_Documentation.xlsx  # Generated documentation
└── README.md
```

## Holiday Credit Rules

The solution includes 36 Holiday Credit Rules covering 9 holidays:
- Christmas Day
- Independence Day
- Labor Day
- Memorial Day
- MLK Day
- New Years Day
- Presidents Day
- Thanksgiving Day
- Veterans Day

### Rule Types

| Type | Count | Description |
|------|-------|-------------|
| Exempt [HOLIDAY] | 9 | Credit rules for exempt employees |
| Exempt [HOLIDAY] FORFEITED | 9 | Forfeit rules for exempt employees |
| Non Exempt [HOLIDAY] | 9 | Credit rules for non-exempt employees |
| Non Exempt [HOLIDAY] FORFEITED | 9 | Forfeit rules for non-exempt employees |

### Contributing Shift Rules

Credit rules reference `H-[HOLIDAY] PENDING` Contributing Shift Rules:
- H-CHRISTMAS PENDING
- H-INDEPENDENCE DAY PENDING
- H-LABOR DAY PENDING
- H-MEMORIAL PENDING
- H-MLK PENDING
- H-NEW YEAR PENDING
- H-PRESIDENTS DAY PENDING
- H-THANKSGIVING PENDING
- H-VETERANS DAY PENDING

## Holiday Profiles

| Profile Name | Description |
|--------------|-------------|
| All Holidays Non Exempt FT | Full-time non-exempt employees |
| All Holidays Non Exempt PT | Part-time non-exempt employees |
| DCNA CSS FT | DCNA CSS full-time |
| DCNA CSS PT | DCNA CSS part-time |
| HSC SEIU | HSC SEIU union employees |
| Non Exempt | Generic non-exempt |
| Non Exempt CPA | Non-exempt CPA |

## Documentation Generator

The `generate_documentation.py` script parses all JSON exports and creates an Excel workbook with:

- Summary sheet with Table of Contents
- Individual sheets for each object type
- Anomaly detection and findings report

### Usage

```bash
python3 generate_documentation.py
```

### Requirements

```bash
pip install pandas openpyxl
```

## Known Issues / Findings

The anomaly detection identifies the following items for review:

1. **Holiday export incomplete** - Missing standard (non-ADVS) holidays for Independence Day, Labor Day, MLK Day, Memorial Day, Presidents Day, Veterans Day
2. **LABOR DAY rules** - Have inconsistent Scheduled Shift Check settings between Exempt and Non-Exempt
3. **All Holidays Non Exempt PT profile** - Veterans Day uses 'Non Exempt' credit rule instead of PT variant
4. **12 holidays have fewer than 3 dates** - Mostly ADVS-prefixed holidays with limited date definitions

## Fixes Applied

During configuration review, the following corrections were made:

1. **FORFEIT PRESIDENTS DAY Contributing Shift** - Fixed incorrect reference from `H-FORFEIT PENDING MLK` to `H-FORFEIT PENDING PRESIDENTS`
2. **Thanksgiving Plus 30 Balance Cascade** - Fixed date pattern from `HOL-Thanksgiving Day Annual` to `HOL-Thanksgiving Day Plus 30 Annual`
3. **Holiday Credit Rules** - Updated Contributing Shift references from old format (`[HOLIDAY] CREDIT TRIGGER`) to new format (`H-[HOLIDAY] PENDING`)
4. **FORFEITED rules** - Cleared Contributing Shift references (FORFEIT PENDING rules removed from scope)
5. **Holiday Profiles** - Removed 6 profiles with zz-prefixed credit rules (test/placeholder configurations)
6. **Removed out-of-scope rules** - `Non Exempt`, `Non Exempt PT`, `HSC UNION FT`, `Exempt FORFEITED`
