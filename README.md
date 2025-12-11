# SkyPoint Audit Tools

This repository contains PowerShell tooling designed to retrieve, organize, and visualize Unified Audit Logs from Microsoft 365 (Exchange Online) for forensic investigations.

## Overview

The toolkit automates the process of gathering audit logs and presenting them in a human-readable, forensic-focused HTML report. It helps investigators quickly identify security incidents, unauthorized access, and data exfiltration risks.

### Key Components

1.  **`CaseFileCreation.ps1` (Orchestrator)**
    *   The main entry point for the tool.
    *   Manages the case directory structure.
    *   Archives previous session data.
    *   Calls the retrieval script to fetch new logs.
    *   Parses raw JSON logs and generates a comprehensive HTML report.

2.  **`GetUnifiedLogData.ps1` (Data Retrieval)**
    *   Connects to Exchange Online.
    *   Retrieves Unified Audit Logs day-by-day to ensure high completeness.
    *   Exports raw data to JSON files for processing.

## Prerequisites

*   **PowerShell**: Version 5.1 or PowerShell 7+.
*   **Exchange Online Management Module**: Required for log retrieval.
    ```powershell
    Install-Module -Name ExchangeOnlineManagement
    ```
*   **Permissions**: The account running the script must have permissions to run `Search-UnifiedAuditLog` (typically *View-Only Audit Logs* or *Audit Logs* roles in Exchange Online).

## Usage

Run the `CaseFileCreation.ps1` script with the required parameters to start a case investigation.

```powershell
.\CaseFileCreation.ps1 -HaloId "Cust562" -TicketId "Case123" -ActorUpn "target@domain.com" -AdminUpn "admin@domain.com" -StartDate "2025-01-01" -EndDate "2025-01-10"
```

### Parameters

| Parameter | Description | Required |
| :--- | :--- | :--- |
| `-HaloId` | Customer Identifier (used for folder structure). | Yes |
| `-TicketId` | Case or Ticket Number (used for folder structure). | Yes |
| `-ActorUpn` | The User Principal Name (email) of the target user to investigate. | Yes |
| `-AdminUpn` | The admin account used to authenticate with Exchange Online. | Yes |
| `-StartDate` | Start date for the audit log search. | Yes |
| `-EndDate` | End date for the audit log search. | Yes |
| `-RecordType` | Filter specific record types (e.g., `SharePointFileOperation`). | No |
| `-Operation` | Filter specific operations (e.g., `FileDeleted`). | No |
| `-SearchText` | Free text search within the audit data. | No |

## Output

The tool generates a self-contained HTML report located in:
`SessionData\<HaloId>\<TicketId>\Report.html`

### Report Features
*   **High-Level Summary**: KPIs, workload breakdown, and top activity.
*   **Timeline View**: Chronological list of all events.
*   **Categorized Sections**: Dedicated views for File Activity, Identity, and Exchange.
*   **Card Layout**: Detailed view of every property in the audit log, ensuring no data is hidden.
*   **Local Time Conversion**: All timestamps are converted from UTC to the local machine time for easier analysis.

## Directory Structure

```text
.
├── CaseFileCreation.ps1    # Main script
├── GetUnifiedLogData.ps1   # Retrieval helper
├── SessionData/            # Generated data storage
│   └── <HaloId>/
│       └── <TicketId>/
│           ├── Audit// filepath: c:\Users\michael.barrett\src\github\SkyPointAuditTools\README.md
```

# SkyPoint Audit Tools

This repository contains PowerShell tooling designed to retrieve, organize, and visualize Unified Audit Logs from Microsoft 365 (Exchange Online) for forensic investigations.

## Overview

The toolkit automates the process of gathering audit logs and presenting them in a human-readable, forensic-focused HTML report. It helps investigators quickly identify security incidents, unauthorized access, and data exfiltration risks.

### Key Components

1.  **`CaseFileCreation.ps1` (Orchestrator)**
    *   The main entry point for the tool.
    *   Manages the case directory structure.
    *   Archives previous session data.
    *   Calls the retrieval script to fetch new logs.
    *   Parses raw JSON logs and generates a comprehensive HTML report.

2.  **`GetUnifiedLogData.ps1` (Data Retrieval)**
    *   Connects to Exchange Online.
    *   Retrieves Unified Audit Logs day-by-day to ensure high completeness.
    *   Exports raw data to JSON files for processing.

## Prerequisites

*   **PowerShell**: Version 5.1 or PowerShell 7+.
*   **Exchange Online Management Module**: Required for log retrieval.
    ```powershell
    Install-Module -Name ExchangeOnlineManagement
    ```
*   **Permissions**: The account running the script must have permissions to run `Search-UnifiedAuditLog` (typically *View-Only Audit Logs* or *Audit Logs* roles in Exchange Online).

## Usage

Run the `CaseFileCreation.ps1` script with the required parameters to start a case investigation.

```powershell
.\CaseFileCreation.ps1 -HaloId "Cust562" -TicketId "Case123" -ActorUpn "target@domain.com" -AdminUpn "admin@domain.com" -StartDate "2025-01-01" -EndDate "2025-01-10"
```

### Parameters

| Parameter | Description | Required |
| :--- | :--- | :--- |
| `-HaloId` | Customer Identifier (used for folder structure). | Yes |
| `-TicketId` | Case or Ticket Number (used for folder structure). | Yes |
| `-ActorUpn` | The User Principal Name (email) of the target user to investigate. | Yes |
| `-AdminUpn` | The admin account used to authenticate with Exchange Online. | Yes |
| `-StartDate` | Start date for the audit log search. | Yes |
| `-EndDate` | End date for the audit log search. | Yes |
| `-RecordType` | Filter specific record types (e.g., `SharePointFileOperation`). | No |
| `-Operation` | Filter specific operations (e.g., `FileDeleted`). | No |
| `-SearchText` | Free text search within the audit data. | No |

## Output

The tool generates a self-contained HTML report located in:
`SessionData\<HaloId>\<TicketId>\Report.html`

### Report Features
*   **High-Level Summary**: KPIs, workload breakdown, and top activity.
*   **Timeline View**: Chronological list of all events.
*   **Categorized Sections**: Dedicated views for File Activity, Identity, and Exchange.
*   **Card Layout**: Detailed view of every property in the audit log, ensuring no data is hidden.
*   **Local Time Conversion**: All timestamps are converted from UTC to the local machine time for easier analysis.

## Directory Structure

```text
.
├── CaseFileCreation.ps1    # Main script
├── GetUnifiedLogData.ps1   # Retrieval helper
├── SessionData/            # Generated data storage
│   └── <HaloId>/
│       └── <TicketId>/
│           ├── Audit