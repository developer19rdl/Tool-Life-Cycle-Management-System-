# Honda Attendance/Inventory Management System (V 1.1.5)

This is the definitive technical documentation for the Honda Inventory application, covering every component's logic, control flows, and SQL queries.

---

## 🏗️ 1. Architecture Overview

- **Framework**: .NET 8.0 Windows Application (WPF).
- **Language**: VB.NET.
- **Database**: MySQL (Centralized Server).
- **Dynamic Table Architecture (DTA)**: Instead of a static schema, the application generates and manages individual SQL tables for every unique Instrument or Gauge Type.
- **Embedded API**: A custom REST server (Port 8080) for external device integration.

---

## 🧠 2. Component Logic & Workflows

### 2.1 Authentication & Session
- **LoginWindow**:
    - **Logic**: Checks for hardcoded admin `Admin / RDL123` first. If failing, queries the `users` table.
    - **Session**: Successfully logged-in username is stored in `Application.Current.Properties("Username")`.
- **AddUser / UserSettings**:
    - Manage standard CRUD for the `users` table. Passwords can be toggled for visibility using a dual `PasswordBox` and `TextBox` sync logic.

### 2.2 Navigation Flow
- **MainWindow**: Acts as the shell with a `MainFrame`.
- **HomePage**: Offers quick access to Instruments and Gauges.
- **CategorySelectionPage**: 
    - Loads all `TypeName` entries from `type_details` filtered by the selected category (Instrument/Gauge).
    - Dynamically generates "Cards" for each type.
    - Navigates to the specific Management Page.

### 2.3 Inventory Management Logic
- **Pagination**: Fixed at 20 records per page using MySQL `LIMIT` and `OFFSET`.
- **Filtering**: Independent dropdowns for Line, Section, Location, and Status. These are populated via `SELECT DISTINCT` queries against the active type table.
- **Control Number Generation (AddEditWindow)**:
    - **Resolution Phase**: App looks up the `GroupCode` from `categorycontrol` (for instruments) or `gauge_categorycontrol` (for gauges) based on the selected `Size`.
    - **Increment Phase**: App queries the highest numeric suffix for that prefix (e.g., `PS57-001` -> `PS57-002`) by casting the substring after the hyphen to an unsigned integer.
- **Import Engine (Excel/CSV)**:
    - Uses a column helper to match headers even if they have extra spaces or dots.
    - Handles bulk insertion using standardized `INSERT` templates.

### 2.4 Lifecycle Management (Write-Off & Reintroduction)
- **Status Tracking**: Uses a `Flag` (0=Active, 1=In-Active) and `InstrumentStatus` string.
- **Document Association**: Both workflows require a PDF upload. The file path is stored in the database, allowing users to open the PDF directly from the **HistoryWindow**.
- **Query Flow**: Write-off inserts into `writeoff` table; Reintroduction deletes-from/updates-status and inserts into `reintroduction` table.

### 2.5 RFID API & Hardware Logic
- **RFIDApiServer**:
    - Spawns a background listener.
    - **Lookup**: When an EPC (RFID) is received, it builds a massive `UNION ALL` query using every table listed in `type_details`.
- **CommManager**:
    - Communicates with USB-Serial devices.
    - Sends Hex commands (e.g., `BB 00 22...`) to triggers scans.
    - Normalizes responses by stripping whitespace and converting to uppercase.

---

## 🗄️ 3. SQL Query Encyclopedia

### 3.1 Master Data & Schema
| Function | Query |
| :--- | :--- |
| Load All Types | `SELECT * FROM type_details ORDER BY Category, TypeName` |
| Get Base Prefix | `SELECT BasePrefix FROM type_details WHERE TypeName = @typeName LIMIT 1` |
| Get Departments | `SELECT * FROM departmentmaster ORDER BY DepartmentName ASC` |
| Add Type | `INSERT INTO type_details (Category, TypeName, PrefixMode, BasePrefix, SerialDigits, TypeImage) VALUES (@cat, @name, @mode, @pre, @dig, @img)` |
| Create Instrument Table | `CREATE TABLE IF NOT EXISTS {tableName} (ID INT AUTO_INCREMENT PRIMARY KEY, Date DATE, Time TIME, InstrumentName VARCHAR(255), InstrumentDescription TEXT, ControlNo VARCHAR(100) UNIQUE, Color VARCHAR(100), InstrumentModelNo VARCHAR(255), Line VARCHAR(100), AssetNo VARCHAR(100), Size VARCHAR(100), MakerName VARCHAR(255), Section VARCHAR(255), Location VARCHAR(255), Remark TEXT, RequestNo VARCHAR(100), CategoryControl VARCHAR(100), RFID_tag VARCHAR(255), Flag INT DEFAULT 0, InstrumentStatus VARCHAR(50) DEFAULT 'Active')` |
| Create Gauge Table | `CREATE TABLE IF NOT EXISTS {tableName} (ID INT AUTO_INCREMENT PRIMARY KEY, Date DATE, Time TIME, GaugeName VARCHAR(255), GaugeDescription TEXT, ControlNo VARCHAR(100) UNIQUE, Color VARCHAR(100), DrgNo VARCHAR(100), MakerName VARCHAR(255), Model VARCHAR(255), Line VARCHAR(100), Size VARCHAR(100), Tol VARCHAR(100), Tol2 VARCHAR(100), Tol3 VARCHAR(100), Section VARCHAR(255), Location VARCHAR(255), DateofAddition VARCHAR(50), RequestNo VARCHAR(100), Remark TEXT, CategoryControl VARCHAR(100), RFID_tag VARCHAR(255), Flag INT DEFAULT 0, InstrumentStatus VARCHAR(50) DEFAULT 'Active')` |
| Insert Instrument | `INSERT INTO {tableName} (Date, Time, InstrumentName, InstrumentDescription, ControlNo, Color, InstrumentModelNo, Line, AssetNo, Size, MakerName, Section, Location, Remark, RequestNo, CategoryControl, RFID_tag, Flag, InstrumentStatus) VALUES (@date, @time, @name, @desc, @control, @color, @model, @line, @asset, @size, @maker, @sec, @loc, @rem, @req, @cat, @rfid, @flag, 'Active')` |
| Insert Gauge | `INSERT INTO {tableName} (Date, Time, GaugeName, GaugeDescription, ControlNo, Color, DrgNo, MakerName, Model, Line, Size, Tol, Tol2, Tol3, Section, Location, DateofAddition, RequestNo, Remark, CategoryControl, RFID_tag, Flag, InstrumentStatus) VALUES (@date, @time, @name, @desc, @control, @color, @drg, @maker, @model, @line, @size, @tol, @tol2, @tol3, @sec, @loc, @dateAdd, @req, @rem, @cat, @rfid, @flag, 'Active')` |
| Update Instrument | `UPDATE {tableName} SET InstrumentName=@name, InstrumentDescription=@desc, ControlNo=@control, Color=@color, InstrumentModelNo=@model, Line=@line, AssetNo=@asset, Size=@size, MakerName=@maker, Section=@sec, Location=@loc, Remark=@rem, RequestNo=@req, CategoryControl=@cat, RFID_tag=@rfid, Flag=@flag [, Date=@date] WHERE ID=@id` |
| Update Gauge | `UPDATE {tableName} SET GaugeName=@name, GaugeDescription=@desc, ControlNo=@control, Color=@color, DrgNo=@drg, MakerName=@maker, Model=@model, Line=@line, Size=@size, Tol=@tol, Tol2=@tol2, Tol3=@tol3, Section=@sec, Location=@loc, DateofAddition=@dateAdd, RequestNo=@req, Remark=@rem, CategoryControl=@cat, RFID_tag=@rfid, Flag=@flag [, Date=@date] WHERE ID=@id` |
| Delete Record | `DELETE FROM {tableName} WHERE ID = @id` |

### 3.2 Inventory Operations
| Function | Query |
| :--- | :--- |
| Paged Read | `SELECT * FROM {tableName} WHERE {filters} ORDER BY Date DESC, Time DESC LIMIT {limit} OFFSET {offset}` |
| Total Count | `SELECT COUNT(*) FROM {tableName} WHERE {filters}` |
| Duplicate Check | `SELECT COUNT(1) FROM {tableName} WHERE ControlNo = @controlNo` |
| Update Status | `UPDATE {tableName} SET Flag = @flag, InstrumentStatus = @status WHERE ControlNo = @controlNo` |
| Search All Tables | `SELECT * FROM ({Dynamic Union of All Tables}) as Combined WHERE ControlNo LIKE '%keyword%'` |

### 3.3 Lifecycle & History
| Function | Query |
| :--- | :--- |
| Insert Write-Off | `INSERT INTO writeoff (WriteOffDate, Time, Type, InstrumentName, ControlNo, Quantity, Line, Section, Color, Reason, Action, ActionNG, DocumentPath, WriteOffNo, RecievedDate, InstrumentStatus, RaisedBy, Interupdated) VALUES (...)` |
| Insert Reintroduction | `INSERT INTO reintroduction (ReintroductionDate, Time, Type, InstrumentName, ControlNo, Quantity, Line, Section, Color, Reason, WriteOffNo, ReintroductionNo, FoundDate, MissingDate, UserDeclaration, InterStatus, InterComent, InterUpdated, CalibrationResult, RaisedBy, Remarks, DocumentPath) VALUES (...)` |
| Get Lifecycle History | `SELECT WriteOffDate AS 'Date', 'WRITE_OFF' AS Action, Reason AS Remarks, DocumentPath FROM writeoff WHERE ControlNo = @ctrl UNION ALL SELECT ReintroductionDate AS 'Date', 'REINTRODUCTION' AS Action, Reason AS Remarks, DocumentPath FROM reintroduction WHERE ControlNo = @ctrl ORDER BY Date DESC` |

### 3.4 User & Security
| Function | Query |
| :--- | :--- |
| Login Check | `SELECT COUNT(*) FROM users WHERE Username=@u AND Password=@p` |
| List Users | `SELECT ID, Username AS UserName, Password, Email, Phone FROM users` |
| Add User | `INSERT INTO users (Username, Password, Email, Phone) VALUES (@u, @p, @e, @ph)` |

### 3.5 RFID Integration (API Internal)
| Function | Query |
| :--- | :--- |
| API Multi-Table Search | `SELECT ControlNo, 'Instrument' AS Type, InstrumentName AS Name, Color, RFID_tag FROM {tbl} WHERE UPPER(REPLACE(RFID_tag, ' ', '')) = @rfid` (Unioned across all tables) |

---

## 🛠️ 4. Technical Specifications
- **Port**: 8080 (REST API).
- **Control Mask**: Automatic Control Numbers follow the pattern `[Prefix]-[XXX]` where `XXX` is a zero-padded 3-digit number.
- **Image Storage**: Type images are stored as **Base64** strings in the `TypeImage` column (LONGTEXT) of `type_details`.
