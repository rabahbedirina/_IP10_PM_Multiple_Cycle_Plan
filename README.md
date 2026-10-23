# SAP PM IP10 Multiple Cycle Plan Automation

**Current Version:** `v1.0.0`

**Author:** Rabah BEDIRINA

**GitHub:** [rabahbedirina](https://github.com/rabahbedirina/)

This project contains a VBScript automation utility designed to handle complex or multiple cycle maintenance plans within SAP transaction `IP10` under **SAP Plant Maintenance (PM)**.

---

## What is SAP Transaction IP10 (Multiple Cycle)?

**IP10** is used to maintain and schedule maintenance plans. While single-cycle plans trigger schedules uniformly, strategy-based or multiple-cycle plans involve specific maintenance packages or sequences. 

This automation reads input parameters row-by-row from a CSV file, navigates through the multi-row table or sequence view using precise keyboard navigation (`Sequence` shifts), selects the target cycle, assigns the requested start date and hour (`07`), and locks in the schedule by saving.

---

## Files Included

- **`_IP10_PM_Multiple_Cycle_Plan.vbs`** — Main VBScript automation script for SAP GUI navigation.
- **`DATA__IP10_Multiple_Cycle_Plan.csv`** (supports fallback filename `DATA_IP10_Multiple_Cycle_Plan.csv`) — Input data file containing plans, start dates, and sequence positions.
- **`README.md`** — Project documentation.

---

## What the Script Does

1. Verifies and loads the CSV data file from the script's current working directory (handling double-underscore naming conventions as a fallback).
2. Skips the CSV header row.
3. Reads each line and splits the semicolon-delimited values into:
   - `MaintenancePlan`
   - `StartDate`
   - `Sequence`
4. Activates the active SAP GUI window.
5. Invokes transaction `/nIP10`.
6. Enters the maintenance plan number and opens the schedule.
7. Triggers the package/sequence view (`Ctrl + F7`).
8. Dynamically calculates and executes `TAB` or `Shift + TAB` offsets based on the provided `Sequence` value to locate the exact target line.
9. Selects the row, inserts the target start date, inputs the hour `07`, and confirms the dialog sequence.
10. Saves the transaction using `Ctrl + S`.
11. Repeats the workflow for all remaining rows and displays a completion summary pop-up.

---

## CSV Data File Format (`DATA__IP10_Multiple_Cycle_Plan.csv`)

The CSV file must use a semicolon (`;`) separator with the following exact column order:

```csv
MaintenancePlan;StartDate;Sequence
12345678;06.08.2026;4
12345678;30.07.2026;8
12345678;25.06.2026;8
12345678;01.08.2026;1
```

### Field Definitions:
- **`MaintenancePlan`**: SAP maintenance plan number (e.g., `12345678`).
- **`StartDate`**: Planned start date to insert into the SAP screen, formatted as `DD.MM.YYYY`.
- **`Sequence`**: Numeric index position used by the script's navigation loop to land on the correct row/package option in the SAP strategy table.

---

## Prerequisites & Requirements

1. **SAP GUI Installed:** Ensure the SAP Logon session is running and accessible on your desktop.
2. **Active SAP Window:** The SAP window title must contain `"SAP"` so `AppActivate` can focus it successfully.
3. **File Placement:** Both the `.vbs` script and the CSV data file must reside in the exact same folder.

---

## How to Run It

1. Open your SAP GUI session and log into your target client.
2. Ensure the SAP window is visible and active.
3. Place `_IP10_PM_Multiple_Cycle_Plan.vbs` and your CSV file in the same directory.
4. Double-click the `.vbs` file to launch the automation.
5. Click **OK** on the startup confirmation prompt and **do not touch your mouse or keyboard** while the script executes.
6. A final pop-up notification will state when all rows have been processed.

---

## Important Behavior & Limitations

- **SendKeys Dependency:** The automation relies on `WScript.Shell.SendKeys` and fixed millisecond delays (`WScript.Sleep`). It requires a stable SAP user interface and standard screen layouts. If layout configurations change, navigation offsets may need fine-tuning.
- **Validation:** Basic trimming is performed, but advanced pre-flight validation of SAP authorization or live plan statuses is not built into the script. Test thoroughly in a development or quality client first.