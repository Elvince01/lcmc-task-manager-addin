# LCMC Task Manager – User Guide

This guide explains how to use the **LCMC Task Manager** Excel add-in pane. You can share this document with colleagues.

---

## 1. Opening the Task Manager pane

1. Open the **TASK MANAGER FORM** workbook in Excel.
2. Go to the **Home** tab on the ribbon.
3. In the **LCMC** group, click **Show LCMC Task Manager**.
4. The **LCMC Task Manager** pane will open on the right side of Excel.

> All task data should be created/edited from this pane rather than typing directly into the sheet.

---

## 2. Layout of the pane

### 2.1 Header

- **Title**: `LCMC Task Manager`.
- **Row**: shows the currently selected Excel row (e.g. `Row: 27`).
- **New Task**: clears the pane so you can prepare a brand new task.
- **Refresh Row**: reloads data from the currently selected row in Excel into the pane.

### 2.2 Company section

- **Company Name** (Column **A**)
  - Free-text field for the client / company name.

### 2.3 Task Details section

- **Task Category** (Column **B**)
  - Multi-select list of categories:
    - Onboarding
    - Amended Product/Rate
    - KYC update
    - Other Task
  - You can tick multiple categories at once.
  - If **Other Task** is selected, the **Other Task** card appears (see below).

- **Assigned to** (Column **C**)
  - Multi-select checklist populated from the **Team members** sheet (column A).
  - You can tick multiple assignees.
  - Ticked assignees are saved as a comma-separated list in Column C (and shown at the top of the list).

- **Broker Lead** (Column **D**)
  - Single-select dropdown populated from the same team member list.

### 2.4 Key Dates section (Columns **E–G**)

All dates are stored in Excel as `d-MMM-yyyy` (e.g. `29-Oct-2025`).

- **Initial Request Date** (Column **E**)
- **Intro Email Date** (Column **F**)
- **⏰ Reminders Starting Date** (Column **G**)

### 2.5 Tasks section (Column **H**)

This section controls the structured multi-line task summary in **Column H**.

#### Task selection

- **Select Tasks**
  - Checkboxes for fixed tasks:
    - Broker Agreement
    - Give-up Agreement
    - Amend Annex A
    - KYC (produced by LCMC)
    - KYC External
  - When you tick a task, its corresponding **task card** becomes visible below.

#### Task cards

For each fixed task (Broker Agreement, Give-up Agreement, etc.) you see:

- **Status**
  - Choice chips: `To Do`, `In Progress`, `Done`.
  - Colour coding (matching the original Google Sheet):
    - To Do: green (`#34a853`)
    - In Progress: orange (`#ff9900`)
    - Done: soft red (`#cc6666`)
- **Comment**
  - Free-text note about that specific task.
- **Deadline**
  - Task-specific due date.
- **Give-up ID** (only on **Give-up Agreement**)
  - Identifier for the give-up agreement.

#### Other Task card

If **Other Task** is selected in **Task Category**, an additional card appears:

- **Task Name**: name of the custom task.
- **Status / Comment / Deadline**: same as for fixed tasks.

> When **Other Task** is selected together with other categories, fixed task cards and the Other Task card can be used together.

### 2.6 Reminders & Notes section (Columns **J–K**)

- **Email Reminders Frequency (EDT)** (Column **J**)
  - Dropdown with predefined options:
    - Every Monday at 8AM
    - Every Monday, Wednesday and Friday at 8AM
    - Every working days at 8AM
    - Every working days at 8AM and 2PM
- **Extra Note** (Column **K**)
  - Free-text area for any final remarks about the overall task.

### 2.7 Audit information

- **History** (read-only)
  - Shows key audit info derived from the tasks (e.g. Task initiated / Status updated / Completed) when available.

### 2.8 Footer actions

- **Add Task to Database**
  - Adds the current pane contents as a **new row** in the first available empty row of the sheet.
- **Save**
  - Writes the current pane contents back to the **currently selected row**.

---

## 3. Column mapping summary

The pane writes to the Excel sheet as follows:

- **A** – Company Name
- **B** – Task Category (comma-separated list if multiple)
- **C** – Assignees (comma-separated list)
- **D** – Broker Lead
- **E** – Initial Request Date
- **F** – Intro Email Date
- **G** – Reminders Starting Date
- **H** – Multi-line task summary (one line per task)
- **J** – Email Reminders Frequency (EDT)
- **K** – Extra Note

**Status colours in Column H**

- The font colour of Column H is set based on the *most urgent* status among its tasks:
  - Any **In Progress** → orange `#ff9900`
  - Else any **To Do** → green `#34a853`
  - Else if all **Done** → red `#cc6666`

> Excel’s API doesn’t support per-line colours inside one cell in this add-in, so the entire cell uses one colour based on the highest-priority status.

---

## 4. Typical workflows

### 4.1 Create a new task

1. In Excel, select any row close to where you want to add the task (for context only).
2. In the pane header, click **New Task**.
3. Fill in:
   - **Company Name**
   - **Task Category** (one or many)
   - **Assigned to** (tick one or more assignees)
   - **Broker Lead** (optional)
   - **Key Dates** (Initial Request, Intro Email, Reminders Starting Date)
4. In **Tasks**:
   - Tick the relevant tasks in **Select Tasks**.
   - For each visible task card, set **Status**, add **Comment** and **Deadline**.
   - If applicable, fill **Give-up ID** and/or use **Other Task**.
5. In **Reminders & Notes**:
   - Choose an **Email Reminders Frequency (EDT)**.
   - Add any **Extra Note**.
6. Click **Add Task to Database**.
   - The add-in finds the first empty row and writes all data (Columns A–H, J–K).
   - That row is then selected in Excel and shown in the **Row** indicator.

### 4.2 Update an existing task

1. In Excel, click any cell in the existing task’s row.
2. In the pane header, click **Refresh Row**.
   - The pane loads all values from that row.
3. Adjust any fields you need: company, categories, assignees, dates, task statuses, comments, deadlines, reminders, extra note.
4. Click **Save**.
   - The current row is updated in place.

---

## 5. KYCai Assistant integration

- For the **KYC (produced by LCMC)** task card, there is a button **Open KYCai Assistant**.
- Clicking it opens the internal KYCai Assistant at:
  - `http://192.168.1.141:8080/`

---

## 6. Notes & limitations

- Always use the **LCMC Task Manager pane** to edit tasks. Manual edits directly in the sheet (especially Column H) may not round-trip perfectly back into the pane.
- Per-line colouring in Column H (different colours for each task line) isn’t supported in this Office add-in version; the whole cell uses one colour based on the most urgent status.
- If you don’t see the latest code/UI changes, you may need to:
  - Close and reopen Excel, and
  - Reload the add-in (depending on how it’s deployed in your environment).
