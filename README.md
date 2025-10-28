# 🐾 SQL Performance Observation Tool (SPOT)

**Version:** v0.9.0-preview  
**Author:** Jake Morgan / Blackcat Data Solutions

---

### Overview

**SPOT** is a lightweight SQL Server performance dashboard for capturing and exploring snapshots of live SQL activity.  
It uses `sp_WhoIsActive`, wait stats, blocking chains, health checks, and AG (Availability Group) health to build an interactive time-series view of what your instance was doing at any given moment.

Snapshots are stored in the `utility.SPOT.*` tables and can later be reloaded or exported for offline analysis.  
The PowerShell GUI (`SPOT.ps1`) provides a dark-themed, tabbed interface with per-sample navigation, execution plan viewing, and connection profile management.

---

### 🧱 Prerequisites

1. **Utility database** — the target database that stores SPOT objects and data.  
2. **`sp_WhoIsActive` (v11.32)** installed in the Utility database.  
   - Download from [https://whoisactive.com](https://whoisactive.com)
3. **PowerShell 5.1+** (Windows only)
4. **SQL Server login** with rights to:
   - Execute `sp_WhoIsActive`
   - Create tables, indexes, and procedures in the Utility DB
   - Capture and query performance data

---

### ⚙️ Installation

1. **Create SPOT schema objects**

   Run the SQL script **`SPOT.sql`** against your *Utility* database.  
   This creates:
   - Tables: `SPOT.Snapshots`, `SPOT.Samples`, `SPOT.WhoIsActive`, `SPOT.WaitStats`, `SPOT.Blocking`, `SPOT.HealthChecks`, `SPOT.AGHealth`
   - Stored procedures: `SPOT.CaptureSnapshot`, supporting metadata, and indexes.

2. **Deploy the GUI client**

   On any workstation or jump box with SQL connectivity to a SQL2012 or higher instance:
   ```powershell
   deploy_SPOT.ps1
   ```
   This will copy the `SPOT.ps1` client files into  
   `Documents\SQL Performance Observation Tool` and set up initial session storage.

3. *(Optional)*  
   Configure a **SQL Agent job** to automatically capture long-running sessions, e.g.:

   ```sql
   IF EXISTS (
       SELECT 1
       FROM sys.dm_exec_requests
       WHERE total_elapsed_time > (60 * 1000) -- > 60 seconds
   )
   BEGIN
       EXEC SPOT.CaptureSnapshot @GetPlans = 1;  -- ensure @GetPlans is always 1 or WIA will not be captured correctly (_slight bug, may be fixed in later releases!_)
   END
   ```
   Each time this runs, a new performance snapshot is captured into the SPOT tables.

---

### 🖥️ Instructions for Use

1. **Launch the GUI**
   ```powershell
   C:\Users\<You>\Documents\SQL Performance Observation Tool\SPOT.ps1
   ```
   This can also be launched from your desktop shortcut (right click `Launch-SPOT.ps1` and run with PowerShell)

2. **Create or select a Session**
   - Click **Save** to store connection details (Server, Username, Password).
   - SPOT keeps sessions in `Sessions.xml` so you can easily switch between environments.  **Passwords are stored in a secure string**.

3. **Test the connection**
   - Use the **Test** button to verify connectivity (_connection tests are initiated automatically when sessions are loaded_)
   - Status turns **green** when successful.

4. **Set capture parameters**
   - Adjust **Sample Interval (seconds)** and **Number of Samples** using the sliders.
   - Typical: 5–10 seconds interval × 5 samples.

5. **Start a capture**
   - Click **Start Capture** to run `SPOT` and collect live data.
   - Estimated runtime = Interval × Samples.
   - SPOT automatically fetches WhoIsActive output, waits, blocking, AG health, and health checks.

6. **Browse results**
   - Use the **timeline slider** to navigate between samples.
   - Each tab (WhoIsActive, WaitStats, Blocking, HealthChecks, AGHealth) shows per-sample data.
   - Red-highlighted rows indicate potential blocking, long waits or sessions of interest.

7. **Save or load snapshots**
   - **Save Snapshot to File** exports the data as JSON for offline use.
   - **Load Snapshot from File** reopens an existing capture.
   - **Load Snapshot from DB** queries historical runs directly from SQL.

8. **View execution plans**
   - Right-click a WhoIsActive row → **View Execution Plan** to open `.sqlplan` in SSMS.

9. **View metadata**
   - Help → **About** for version, author, and GitHub link.
   - Help → **How to Guide** shows a quick usage summary.

---

### 🧩 Object Overview

| Object | Type | Purpose |
|--------|------|----------|
| `SPOT.Snapshots` | Table | Stores snapshot headers |
| `SPOT.Samples` | Table | Each capture iteration (timestamped) |
| `SPOT.WhoIsActive` | Table | WhoIsActive output per sample |
| `SPOT.WaitStats` | Table | Wait stats per sample |
| `SPOT.Blocking` | Table | Blocking relationships |
| `SPOT.HealthChecks` | Table | Health check results |
| `SPOT.AGHealth` | Table | Availability Group sync info |
| `SPOT.CaptureSnapshot` | Proc | Performs multi-sample capture run |

---

### 📦 Files Included

| File | Description |
|------|-------------|
| `SPOT.sql` | Core SQL schema and stored procedure definitions |
| `SPOT.ps1` | GUI dashboard client |
| `deploy_SPOT.ps1` | Installation/deployment helper |
| `Sessions.xml` | Auto-generated session profiles (per-user) |

---

### 🧠 Notes

- The GUI auto-launches in **STA mode** if needed for WinForms.
- You can run multiple SPOT captures against different servers\instances concurrently.
- Captured snapshots remain queryable via T-SQL in the Utility database.
- Execution plans are saved as XML and opened automatically in SSMS/Plan Explorer/your default .sqlplan viewer.

---

### 🐙 Source & Issues

Source code and updates:  
[https://github.com/jakemorgangit/SPOT](https://github.com/jakemorgangit/SPOT)

Bug reports and feature requests welcome!

---

### 📄 Licence

MIT Licence  
© 2025 Jake Morgan / Blackcat Data Solutions
