# AuditUnifiedLogCaseBuilder Copilot Instructions

This repository contains PowerShell tooling to retrieve, organize, and visualize Unified Audit Logs from Microsoft 365 (Exchange Online) for forensic cases.

## Architecture & Data Flow

1.  **Orchestration (`CaseFileCreation.ps1`)**:
    -   **Role**: Main entry point. Orchestrates cleanup, retrieval, and reporting.
    -   **Parameters**: `HaloId` (Customer), `TicketId` (Case), `ActorUpn`, `AdminUpn`, Date Range.
    -   **Workflow**:
        1.  **Archive**: Zips existing content in `SessionData/<HaloId>/<TicketId>/` to an archive file (e.g., `Archive_yyyyMMddHHmmss.zip`).
        2.  **Clean**: Removes all non-archive files from the case directory.
        3.  **Retrieve**: Calls `GetUnifiedLogData.ps1` to fetch fresh logs.
        4.  **Report**: Generates the HTML report from the new logs.

2.  **Data Retrieval (`GetUnifiedLogData.ps1`)**:
    -   **Role**: Helper script called by `CaseFileCreation.ps1`.
    -   **Process**: Fetches logs day-by-day using `Search-UnifiedAuditLog`.
    -   **Output**: Stores raw JSON in `SessionData/<HaloId>/<TicketId>/`.

## Developer Workflow

### Prerequisites
- **PowerShell 5.1+** or PowerShell 7+.
- **ExchangeOnlineManagement Module**: Required for retrieval (`Install-Module -Name ExchangeOnlineManagement`).

### Running the Tools
**Main Execution**:
```powershell
.\CaseFileCreation.ps1 -HaloId "Cust562" -TicketId "Case123" -ActorUpn "target@domain.com" -AdminUpn "admin@domain.com" -StartDate "2025-01-01" -EndDate "2025-01-10"
```

## HTML Report Requirements (`CaseFileCreation.ps1`)

The report must be structured by **investigative intent**, not just raw schema.

### Core Sections
1.  **Case Header**: Case ID, Target User, Time Range, Generation Timestamp.
2.  **High-Level Summary (KPIs)**: Total events, breakdown by workload (SharePoint, Exchange, Entra), top operations, top IPs.
3.  **Timeline View**: Chronological list of actions (Time, Workload, Operation, Description).
4.  **Files (SharePoint/OD)**:
    -   *Mass Activity*: High volume access/delete.
    -   *Deletions*: `FileDeleted`, `FileRecycled`.
    -   *Downloads/Access*: `FileDownloaded`, `FileAccessed`.
    -   *Sharing*: `SharingSet`, `SharingInvitationCreated`.
5.  **Mail (Exchange)**:
    -   *Access*: `MailItemsAccessed`, `MailboxLogin`.
    -   *Actions*: `MoveToDeletedItems`, `SoftDelete`, `HardDelete`.
6.  **Identity (Entra ID)**:
    -   *Sign-ins*: Locations, failures, impossible travel.
    -   *Privilege*: Role changes, group membership changes.
7.  **Correlations**:
    -   *By IP*: Event counts, first/last seen.
    -   *By Client*: UserAgent analysis.
8.  **Anomalies**: Derived findings (e.g., "Mass Delete detected", "After-hours activity").
9.  **Raw Data Appendix**: Full event list for forensic traceability.

### Data Presentation Strategy

#### 1. Grouping Hierarchy
To make the report consumable, group data in this hierarchy:
-   **Level 1: Workload (The "Where")**:
    -   *File Activity*: `SharePointFileOperation`
    -   *Identity & Access*: `AzureActiveDirectoryStsLogon`, `AzureActiveDirectory`
    -   *Exchange*: `ExchangeItem`, `ExchangeAdmin`
    -   *Business Apps*: `CRM`, `PowerPlatformAdministratorActivity`
    -   *Collaboration*: `MicrosoftTeams`
-   **Level 2: Intent (The "What")**:
    -   *Access/Read*: `FileAccessed`, `FilePreviewed`, `UserLoggedIn`, `MailItemsAccessed`
    -   *Modification*: `FileModified`, `Set-Mailbox`, `UpdateUser`
    -   *Exfiltration Risk*: `FileSyncDownloadedFull`, `FileDownloaded`
    -   *Deletion*: `FileDeleted`, `HardDelete`, `SoftDelete`
-   **Level 3: Context (The "Details")**:
    -   *Actor*: `UserIds`
    -   *Target*: `SourceFileName`, `EntityName`, or `ClientIP`
    -   *Time*: `CreationDate`

#### 2. Event Formatting ("Card" Layout)
Do **not** use a generic table for event details, as schemas vary wildly. Use a **Card/Block** layout:
-   **Header**: Standardized fields (`Time`, `Operation`, `User`, `Workload`).
-   **Body**: A **Dynamic Key-Value List** (HTML `<dl>`) of *all* properties found in the `AuditData` JSON.
    -   *Requirement*: Iterate through every property in `AuditData`. Do not hardcode columns.
    -   *Requirement*: Render complex nested objects as JSON strings or nested lists.
    -   *Goal*: Ensure **NO data is hidden** due to schema variability.

### Implementation Notes
-   **Navigation**: Include a top-level `<nav>` with anchors to each major section (1–9).
-   **Defensive Parsing**: Handle missing properties gracefully (not all events have all fields).
-   **Data Source**: Always base the report on the local JSON files in `SessionData`.
-   **Self-Contained**: The output must be a single HTML file with no external dependencies (embed all CSS/JS).

### Script Architecture (Separation of Concerns)
To ensure readability and maintainability, `CaseFileCreation.ps1` must strictly separate concerns:
1.  **Data Layer**: Functions that read JSON and normalize data (e.g., `Import-CaseData`). Returns `PSCustomObject` arrays.
2.  **Logic Layer**: Functions that filter, group, and analyze data (e.g., `Get-Anomalies`). Returns `PSCustomObject` arrays. **NO HTML generation here.**
3.  **Presentation Layer**: Functions that accept objects and return HTML strings (e.g., `New-HtmlTable`).
    -   *Rule*: Do not mix HTML string concatenation inside analysis loops.
    -   *Rule*: Use a `StringBuilder` for the final report assembly.
    -   *Rule*: Embed all styles and scripts inline.

## Error Handling & Logging

-   **Case Log**: All critical steps must log to `Case.log` in the case directory (`SessionData/<HaloId>/<TicketId>/Case.log`).
    -   Include: Timestamp, Severity (INFO/WARN/ERROR), Message.
    -   Scope: EXO errors, JSON parse failures, anomaly detection counts.
-   **Exit Codes**:
    -   `0`: Success.
    -   `1`: Fatal (e.g., EXO connection failed).
    -   `2`: Warning (e.g., No data found).
    -   `3`: Error (e.g., Critical parse failure).
-   **Graceful Partial Success**:
    -   If retrieval fails mid-stream (e.g., network error after day 3 of 10), do **not** abort.
    -   Build the report with available data.
    -   Mark the report header clearly as **"PARTIAL DATA"**.

## Project Conventions

-   **Directory Structure**: `SessionData/<HaloId>/<TicketId>/` is the case root.
    -   `HaloId` = Customer Number.
    -   `TicketId` = Case/Ticket Number.
-   **Parameter Validation**: Use `[Parameter(Mandatory=$true)]` for critical inputs.
-   **Date Handling**: Use `[DateTime]` types for parameters; `yyyyMMdd` for filenames. **Crucial**: Source logs are UTC; convert to Local Time for all report displays.
-   **JSON Output**: All intermediate data is serialized with `ConvertTo-Json -Depth 6`.

## Critical Files
-   `GetUnifiedLogData.ps1`: Data retrieval logic.
-   `CaseFileCreation.ps1`: HTML report generation logic.
-   `SessionData/`: Data storage directory.
