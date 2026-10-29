# RM Tracker Automation

This Windows application combines the `Franklinton`, `Blair`, and `Salem`
sheets from a purchase-order tracker, creates an analysis workbook, identifies
missing order confirmations and delivery updates, maps supplier email
addresses, and optionally saves follow-up drafts in Classic Outlook.

Private tracker workbooks, supplier contact lists, generated reports, audit
evidence, and packaged executables are intentionally excluded from this
repository. Supply your own workbooks when running the application.

## Requirements

- Windows and Python 3.12 or a compatible Python 3 release
- Microsoft Excel for opening the generated workbook
- Classic Outlook with a configured mail profile for direct draft creation

Install the dependencies from the project directory:


```powershell
python -m pip install -r requirements.txt
```

## Run

Start the graphical interface:

```powershell
python rm_tracker_app.py
```

Or provide the files from PowerShell:

```powershell
$tracker = "C:\path\to\RM_Tracker.xlsx"
$contacts = "C:\path\to\Vendor_Contact_List.xlsx"
$output = "C:\path\to\RM_Tracker_Pivot_Analysis_Final.xlsx"

python rm_tracker_app.py `
  --tracker $tracker `
  --contacts $contacts `
  --output $output
```

Add `--drafts` to create Outlook drafts. The application saves drafts but does not send them.

## Test

Run the test suite from the project root:

```powershell
python -m unittest discover -s tests -v
```

The main audit regression suite uses private, locally retained workbook
fixtures and is skipped when those fixtures are unavailable. The synthetic
Monthly Analysis tests remain portable.

## Build

Install PyInstaller, then use the packaging specification:

```powershell
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --clean rm_tracker_app_pivot.spec
```

The generated executable is written to the `dist` directory. The packaged
binary is not committed to source control; publish it separately as a GitHub
Release when required.
