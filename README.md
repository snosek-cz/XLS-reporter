# XLS Reporter

Automated monthly time-tracking report generator for consulting teams.
Takes a **Details** sheet (raw time logs from your time-tracking tool) and generates a clean **Overview** sheet aggregating hours by person and week — then optionally consolidates all months into an annual workbook.

---

## How It Works

```
Details_April_2026.xlsx          →       Details_April_2026.xlsx
─────────────────────────               ─────────────────────────
│ Details April 2026 │           →      │ April 2026   (NEW) │  ← overview generated
│ (raw time logs)    │                  │ Details April 2026 │  ← untouched
─────────────────────────               ─────────────────────────
```

The **Overview sheet** aggregates daily time log entries into weekly summaries per person and project:

| Resource Name      | POD Name   | WBS Code         | Rate | Week 1 | Week 2 | Week 3 | Week 4 | Total Hours | Fee     |
|--------------------|------------|------------------|------|--------|--------|--------|--------|-------------|---------|
| Edmund Blackadder  | Cobra POD  | GBL0007741.2.11  | 185  | 4      | 19     | 26     | 19     | =SUM(...)   | =K*E    |
| S. Baldrick        | Cobra POD  | GBL0007741.2.11  | 95   | 22     | 40     | 40     | 38     | =SUM(...)   | =K*E    |
| ...                |            |                  |      |        |        |        |        |             |         |
| **TOTAL**          |            |                  |      | =SUM   | =SUM   | =SUM   | =SUM   | =SUM        | =SUM    |

---

## Repository Structure

```
XLS-reporter/
├── generate_overview.py        Core script: reads Details sheet, writes Overview sheet
├── consolidate_annual.py       Consolidates monthly files into Annual_YEAR.xlsx
├── generate_test_data.py       Generates 12 months of dummy test data
├── Hours_report_template.xlsx  Reference template (Oct + Dec 2025)
├── macos_setup/
│   ├── setup_macos.sh          One-time installer for macOS automation
│   ├── uninstall_macos.sh      Clean removal of all installed components
│   ├── folder_action_trigger.sh  Folder Action pipeline script
│   └── quick_action_run.sh     Finder right-click Quick Action script
└── test_data/
    └── Details_*.xlsx          12 months of sample data (Apr 2026 – Mar 2027)
```

---

## Quick Start — macOS Automated Setup

### 1. Clone the repo
```bash
git clone https://github.com/snosek-cz/XLS-reporter.git
cd XLS-reporter
```

### 2. Run the installer
```bash
bash macos_setup/setup_macos.sh
```

The setup will ask you:
```
Parent folder location [~/Documents]:
Inbox folder name [XLS-Inbox]:
Done folder name [XLS-Done]:
```

It will then:
- Create your **Inbox** and **Done** folders
- Install an isolated Python environment (no system packages touched)
- Attach a **Folder Action** to your Inbox folder
- Install a **Quick Action** in Finder (right-click menu)

### 3. Enable Folder Actions (macOS)
1. Right-click your Desktop → **Services → Folder Actions Setup…**
2. Tick **Enable Folder Actions** ✅
3. Find your Inbox folder in the left panel — it should be listed with **XLS-Reporter-Folder-Action** checked

---

## Usage

### Automatic (Folder Action)
Drop any `Details_*.xlsx` file into your **Inbox** folder:
```
~/Documents/XLS-Inbox/Details_May_2026.xlsx
```
- Overview sheet is generated inside the file
- File moves to your **Done** folder automatically
- `Annual_2026.xlsx` is updated in the Done folder
- macOS notification confirms completion

### Manual — Right-click
Right-click any `Details_*.xlsx` in Finder → **Quick Actions → Generate XLS Overview**

### Manual — Terminal
```bash
# Generate overview for a single file
~/Library/Scripts/XLS-Reporter/venv/bin/python3 \
  ~/Library/Scripts/XLS-Reporter/generate_overview.py \
  /path/to/Details_April_2026.xlsx

# Consolidate all months in a folder into Annual_YEAR.xlsx
~/Library/Scripts/XLS-Reporter/venv/bin/python3 \
  ~/Library/Scripts/XLS-Reporter/consolidate_annual.py \
  ~/Documents/XLS-Done/
```

---

## Input File Format — Details Sheet

The Details sheet must be named `Details MONTH YEAR` (e.g. `Details April 2026`).

| Column | Field                  | Notes                        |
|--------|------------------------|------------------------------|
| A      | Client                 |                              |
| B      | Name                   | Resource/person name         |
| C      | Date                   | Format: `dd.mm.yyyy`         |
| D      | Project                | = POD name in overview       |
| E      | WBS                    | Work package code            |
| F      | Project role           |                              |
| G      | Feature                |                              |
| H      | PBI                    |                              |
| I      | Description            |                              |
| J      | Logged time (hours)    | Numeric hours per day        |
| K      | Rate                   | Hourly rate                  |
| L      | Rate per MD/per hour   | `per hour` text              |
| M      | Total fee              | = J × K                      |
| N      | Week num               | `=ISOWEEKNUM(C)` formula     |

---

## Week Mapping Logic

ISO week numbers from the date column are sorted by their first occurrence within the month and mapped to **Week 1, Week 2, …** sequentially.

> Example — December 2025: ISO week 49 → Week 1, ISO week 50 → Week 2, …, ISO week 1 (Jan) → Week 5

Months with 4 or 5 weeks are handled automatically.

---

## Uninstall

```bash
bash macos_setup/uninstall_macos.sh
```

Removes all installed scripts, the Folder Action and Quick Action. Your processed files in the Done folder are **always preserved**.

---

## Troubleshooting

### Folder Action not firing
1. Check Folder Actions are enabled: right-click Desktop → **Services → Folder Actions Setup** → tick **Enable Folder Actions**
2. Check the log: `tail -f ~/Library/Logs/XLS-Reporter.log`
3. Try the manual Terminal command above to verify the Python script works
4. On macOS Sonoma/Sequoia: **System Settings → Privacy & Security → Automation** — ensure Terminal/System Events has permission

### `externally-managed-environment` error
The setup uses an isolated venv — this error should not occur. If it does, re-run `bash macos_setup/setup_macos.sh`.

### Overview sheet not matching expected values
Verify the Details sheet name starts with `Details ` (with a space) and the date column uses datetime values.

---

## Test Data

Generate 12 months of sample data (April 2026 – March 2027):
```bash
python3 generate_test_data.py
```

The test team is generously staffed from the Blackadder universe, working diligently for Deloitte UK:

| Name | Role | POD | Rate |
|------|------|-----|------|
| Edmund Blackadder | Senior Manager | Cobra POD | £185/h |
| S. Baldrick | Junior Developer | Cobra POD | £95/h |
| George St. Barleigh | Senior Developer | Cobra POD | £130/h |
| Nurse Darling | UX/UI Lead | Viper POD | £140/h |
| Gen. Melchett | Cloud Architect | Viper POD | £210/h |
| Capt. Kevin Darling | QA Test Lead | Viper POD | £120/h |
| Bobbie Parkhurst | Full Stack Developer | Python POD | £125/h |
| Doris Miggins | Data Engineer | Python POD | £135/h |
| Lord Percy Percy | DevOps Engineer | Python POD | £145/h |
| Mrs. Miggins | Business Analyst | Cobra POD | £115/h |

---

## Requirements

- **macOS** (for Folder Action and Quick Action automation)
- **Python 3.8+** (setup creates an isolated venv automatically)
- **openpyxl** (installed automatically in the venv)
- Input files must be `.xlsx` format (not legacy `.xls`)
