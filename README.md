# Axis Max Life Insurance Portal - Selenium Automation

Automated end-to-end testing framework for the Axis Max Life Insurance premium calculator portal. This tool processes test data from Excel, fills multi-step insurance forms, downloads Benefit Illustration PDFs, extracts policy details, and generates comprehensive test reports.

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Project Structure](#project-structure)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Test Data Format](#test-data-format)
- [Output](#output)
- [Portal Flow](#portal-flow)
- [Troubleshooting](#troubleshooting)
- [Recent Fixes](#recent-fixes)
- [Contributing](#contributing)

## 🎯 Overview

This automation framework streamlines the testing of the Axis Max Life Insurance premium calculator by:
- Reading test data from Excel spreadsheets
- Automating form filling across multiple portal pages
- Downloading and parsing Benefit Illustration PDFs
- Extracting and validating Equote Numbers and Premium amounts
- Generating detailed test reports in CSV and Excel formats

## ✨ Features

### Core Functionality
- **Multi-Row Test Execution**: Process up to 6 test cases in a single run
- **Intelligent Form Filling**: Handles complex multi-step forms with dynamic field mapping
- **PDF Processing**: Automatic download and data extraction from Benefit Illustration PDFs
- **Data Validation**: Compares portal values with PDF values for accuracy
- **Comprehensive Reporting**: Generates both CSV and Excel reports with test results
- **Screenshot Capture**: Automatic screenshots on errors for debugging
- **Robust Error Handling**: Retry logic and fallback strategies for reliability

### Advanced Features
- **Smart Field Mapping**: 
  - Income bracket conversion (numeric to range)
  - Life cover parsing (supports "3 crore", "50 lakh", numeric formats)
  - Occupation and education normalization
- **Browser State Management**: Clean state between test cases
- **Multiple Click Strategies**: XPath, CSS selectors, and JavaScript fallbacks
- **Configurable Waits**: Adjustable timeouts for different page loads
- **Headless Mode Support**: Run with or without visible browser

## 📁 Project Structure

```
watermellon/
├── axis_max_life_automation.py    # Main automation script
├── check_excel.py                 # Excel data validation utility
├── README.md                      # This file
├── test data/
│   └── insurance_test_data.xlsx  # Test data spreadsheet
├── downloads/                     # PDF downloads directory
│   ├── BI-7762SKXG.pdf
│   └── BI-8250WXFI.pdf
├── plan&automation_flow/                   
│   └── implementation_plan.md         # Technical implementation details
└── test_results/                  # Generated reports
    ├── report_20260511_153842.csv
    └── report_20260511_153842.xlsx
```

## 🔧 Prerequisites

- **Python**: 3.8 or higher
- **Google Chrome**: Latest version
- **Operating System**: Windows (tested), Linux, or macOS

## 📦 Installation

1. **Clone or download the project**:
   ```bash
   cd watermellon
   ```

2. **Install required Python packages**:
   ```bash
   pip install selenium openpyxl PyPDF2 webdriver-manager
   ```

   **Package Details**:
   - `selenium`: Web browser automation
   - `openpyxl`: Excel file reading/writing
   - `PyPDF2`: PDF text extraction
   - `webdriver-manager`: Automatic ChromeDriver management

3. **Verify installation**:
   ```bash
   python check_excel.py
   ```

## ⚙️ Configuration

### Script Configuration

Edit `axis_max_life_automation.py` to customize behavior:

```python
# Configuration section (lines ~52-60)
HEADLESS = False          # Set True for headless mode
MAX_ROWS = 6              # Number of test rows to process (1-6)
DEFAULT_WAIT = 20         # Seconds for explicit waits
DOWNLOAD_WAIT = 30        # Seconds to wait for PDF download
```

### Portal URL
The script targets the Axis Max Life premium calculator. The URL is configured in the script:
```python
PORTAL_URL = "https://www.axismaxlife.com/term-insurance-plans/premium-calculator..."
```

### Directory Paths
All paths are automatically configured relative to the script location:
- Test Data: `./test data/insurance_test_data.xlsx`
- Downloads: `./downloads/`
- Results: `./test_results/`
- Screenshots: `./screenshots/`

## 🚀 Usage

### Basic Usage

Run the automation script:
```bash
python axis_max_life_automation.py
```

The script will:
1. Load test data from Excel
2. Process the first 6 rows (configurable)
3. Fill forms for each test case
4. Download Benefit Illustration PDFs
5. Extract and validate data
6. Generate reports in `test_results/`

### Headless Mode

For faster execution without visible browser:
```python
# In axis_max_life_automation.py
HEADLESS = True
```

Then run:
```bash
python axis_max_life_automation.py
```

### Processing Specific Rows

To process fewer test cases, modify `MAX_ROWS`:
```python
MAX_ROWS = 3  # Process only first 3 rows
```

## 📊 Test Data Format

### Excel Structure

The test data file (`test data/insurance_test_data.xlsx`) must contain these columns:

| Column Name | Format | Example | Required |
|-------------|--------|---------|----------|
| Full Name | Text | Vikram Banerjee | Yes |
| Date of birth | DD/MM/YYYY | 19/08/1982 | Yes |
| Mobile | 10 digits | 9069041652 | Yes |
| Annual income | Numeric | 3,812,052 | Yes |
| Gender | Male/Female | Male | Yes |
| Occupation | Text | Self Employed | Yes |
| Education | Text | Grad or above | Yes |
| Life cover | Text/Numeric | 3 crore | Yes |
| Cover till age | Numeric | 65 | Yes |
| Critical Illness Rider | Yes/No | No | Yes |
| email id | Email | vikram.banerjee@gmail.com | Yes |
| Pincode | 6 digits | 560016 | Yes |

### Field Mappings

#### Annual Income Brackets
The script automatically maps numeric income to portal brackets:

| Income Range (₹) | Portal Bracket |
|------------------|----------------|
| < 5,00,000 | < 5 |
| 5,00,000 - 7,00,000 | 5 - 7 |
| 7,00,000 - 10,00,000 | 7 - 10 |
| 10,00,000 - 20,00,000 | 10 - 20 |
| > 20,00,000 | >20 |

#### Occupation Mapping
| Excel Value | Portal Value |
|-------------|--------------|
| Salaried | Salaried |
| Self Employed | Self-employed/Business |
| Housewife | Self-employed/Business |
| Business | Self-employed/Business |

#### Education Mapping
| Excel Value | Portal Value |
|-------------|--------------|
| Graduate | Graduate & Above |
| Grad or above | Graduate & Above |
| Post Graduate | Graduate & Above |
| 12th Pass | 12th Pass |
| 10th Pass | 10th Pass |

#### Life Cover Formats
Supports multiple formats:
- **Crore notation**: "3 crore", "1.5 crore"
- **Lakh notation**: "50 lakh", "75 lac"
- **Numeric**: "30000000", "5000000"

## 📈 Output

### Test Reports

Generated in `test_results/` directory with timestamp:

**CSV Format** (`report_YYYYMMDD_HHMMSS.csv`):
```csv
Test Case ID,Name of Insurer,Equote Number,Premium,PDF Match Status
TC1,Vikram Banerjee,BI-7762SKXG,₹15,234,Match
TC2,Manish Menon,BI-8250WXFI,₹18,456,Match
```

**Excel Format** (`report_YYYYMMDD_HHMMSS.xlsx`):
- Same data as CSV
- Formatted headers
- Auto-sized columns

### Report Columns

| Column | Description |
|--------|-------------|
| Test Case ID | Sequential identifier (TC1, TC2, ...) |
| Name of Insurer | Full name from test data |
| Equote Number | Extracted policy quote number |
| Premium | Monthly/annual premium amount |
| PDF Match Status | Match/Mismatch/Not Available |

### Downloaded PDFs

Benefit Illustration PDFs are saved to `downloads/` with naming pattern:
```
BI-{EQUOTE_NUMBER}.pdf
```

Example: `BI-7762SKXG.pdf`

## 🔄 Portal Flow

The automation navigates through these pages:

```
┌─────────────────────────────────────────────────────────────┐
│ 1. Landing Page                                             │
│    - Full Name, DOB, Mobile, Income Bracket                │
│    - NRI Status (always "No")                              │
└─────────────────────────────────────────────────────────────┘
                          ↓
┌─────────────────────────────────────────────────────────────┐
│ 2. Required Details Modal                                   │
│    - Gender, Tobacco (No), Language (English)              │
│    - Occupation, Education                                  │
│    - Diabetic (No), Marital Status (Single)                │
└─────────────────────────────────────────────────────────────┘
                          ↓
┌─────────────────────────────────────────────────────────────┐
│ 3. Step 1/4: Customize Plan                                │
│    - Life Cover Amount                                      │
│    - Cover Till Age                                         │
└─────────────────────────────────────────────────────────────┘
                          ↓
┌─────────────────────────────────────────────────────────────┐
│ 4. Step 2/4: Enhance Plan (Riders)                         │
│    - Critical Illness Rider (Yes/No)                       │
│    - Skip or Proceed                                        │
└─────────────────────────────────────────────────────────────┘
                          ↓
┌─────────────────────────────────────────────────────────────┐
│ 5. Step 3/4: Eligibility & Residential                     │
│    - Email, Annual Income, Pincode, City                   │
│    - Download Benefit Illustration PDF                     │
└─────────────────────────────────────────────────────────────┘
                          ↓
┌─────────────────────────────────────────────────────────────┐
│ 6. Step 4/4: Summary & Equote                              │
│    - Extract Equote Number                                  │
│    - Extract Premium Amount                                 │
└─────────────────────────────────────────────────────────────┘
```

## 🐛 Troubleshooting

### Common Issues

#### 1. ChromeDriver Version Mismatch
**Error**: `SessionNotCreatedException: session not created`

**Solution**: The script uses `webdriver-manager` which auto-updates ChromeDriver. Ensure Chrome browser is up to date:
```bash
# Check Chrome version
chrome --version

# Update webdriver-manager cache
pip install --upgrade webdriver-manager
```

#### 2. Element Not Found
**Error**: `NoSuchElementException` or `TimeoutException`

**Solution**: 
- Check if portal structure has changed
- Increase `DEFAULT_WAIT` timeout
- Review screenshots in `screenshots/` folder
- Enable DEBUG logging for detailed output

#### 3. PDF Download Fails
**Error**: PDF not found in downloads folder

**Solution**:
- Increase `DOWNLOAD_WAIT` timeout
- Check Chrome download settings
- Verify `downloads/` folder permissions
- Disable popup blockers

#### 4. Excel File Not Found
**Error**: `FileNotFoundError: insurance_test_data.xlsx`

**Solution**:
```bash
# Verify file exists
dir "test data\insurance_test_data.xlsx"

# Check file path in script matches your structure
```

#### 5. Multiple Test Cases Fail After First Success
**Issue**: Test case 1 succeeds, but cases 2-6 fail

**Solution**: This was fixed in the latest version. Ensure you have the updated script with:
- Enhanced browser state clearing
- Improved button click strategies
- Navigation verification

See `FIXES_APPLIED.md` for details.

### Debug Mode

Enable detailed logging:
```python
# In axis_max_life_automation.py
logging.basicConfig(
    level=logging.DEBUG,  # Change from INFO to DEBUG
    ...
)
```

### Manual Testing

Test individual components:

**Test Excel Reading**:
```bash
python check_excel.py
```

**Test Life Cover Parsing**:
```bash
python test_parse.py
```

## 🔧 Recent Fixes

### Version 2.0 (May 11, 2026)

Major improvements for multi-row test execution:

1. **MAX_ROWS Configuration**: Changed from 1 to 6 to process all test cases
2. **Browser State Management**: Complete state clearing between test cases
3. **Step 2 Button Clicking**: Enhanced with multiple fallback strategies
4. **Navigation Verification**: Added checks to confirm successful page transitions
5. **Retry Logic**: 3-attempt retry for landing page navigation
6. **Wait Times**: Increased for page transitions and JavaScript execution

**Key Improvements**:
- ✅ All 6 test cases now execute successfully
- ✅ Robust button clicking with 3-tier fallback strategy
- ✅ Enhanced error handling and logging
- ✅ Verification of page state after navigation

See `FIXES_APPLIED.md` for complete details.

## 📝 Logging

The script provides detailed console output:

```
12:34:56  INFO      Loaded 6 test rows
12:34:57  INFO      ═══ Test Case 1/6: Vikram Banerjee ═══
12:34:58  INFO        → Landing page
12:34:59  INFO          Income bracket: >20
12:35:00  INFO        ✓ Landing page submitted
12:35:02  INFO        → Details modal
12:35:03  INFO          Gender: Male
12:35:04  INFO          Tobacco: No
...
```

**Log Levels**:
- `DEBUG`: Detailed click attempts and element searches
- `INFO`: Major steps and successful operations
- `WARNING`: Fallback strategies and retries
- `ERROR`: Failures and exceptions

### Code Modifications

Key areas for customization:

**Field Mappings** (lines ~65-95):
```python
OCCUPATION_MAP = {...}
EDUCATION_MAP = {...}
```

**Timeouts** (lines ~52-60):
```python
DEFAULT_WAIT = 20
DOWNLOAD_WAIT = 30
```

**Portal Flow** (lines ~1400-1600):
- `fill_landing_page()`
- `fill_details_modal()`
- `fill_step1_customize()`
- etc.
  
---

**Last Updated**: May 11, 2026  
**Version**: 2.0  
**Python Version**: 3.8+  
**Tested On**: Windows 10/11, Chrome Latest
