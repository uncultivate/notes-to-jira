# Notes to Jira

Imports Lotus Notes mailbox conversations into Jira Data Center.

- One Jira issue is created for the first email in a conversation.
- Later replies in the same thread are added as comments on that issue.
- Already-processed documents are skipped using a per-mailbox state CSV, so the same email is not imported twice.

There are two scripts:

| Script | Purpose |
| --- | --- |
| `Setup-JiraEmailImporter.ps1` | Store the Jira token and validate config against Jira |
| `Task-NotesQuery.ps1` | Scan Notes and create issues / comments |

## Requirements

- Windows 10/11 or Windows Server
- Windows PowerShell 5.1
- Lotus Notes installed (the import uses 32-bit Notes COM)
- Network access to the Notes/Domino server and to Jira
- A Jira Personal Access Token for an account that can:
  - browse the target project
  - create issues of the configured type
  - set the configured component, labels, due date, and Epic Link
  - add comments

The Windows account that stores the token must be the same account that later runs the import.

## Quick start

### 1. Configure `config.json`

Copy or edit `config.json` in the application folder. Required top-level sections are `modules`, `security`, `jira`, `notes`, `processing`, and `mailboxes`. See [Configuration](#configuration) for every setting.

Do not put the Jira token in `config.json`.

### 2. Store the Jira PAT and validate Jira targets

Setup does not need 32-bit PowerShell. From the application folder:

```powershell
.\Setup-JiraEmailImporter.ps1
```

On first run it prompts for a Jira Personal Access Token and stores it in Windows Credential Manager under the name in `security.credentialTarget` (default: `NotesToJira:JiraPAT`).

It then checks:

- Jira connectivity
- Epic Link field availability
- `jira.defaultEpicKey` is actually an Epic
- each enabled mailbox’s project, issue type, and component exist

To replace an existing token:

```powershell
.\Setup-JiraEmailImporter.ps1 -ReplacePat
```

### 3. Preview the Notes import (32-bit PowerShell)

The import **must** run in 32-bit Windows PowerShell because Lotus Notes COM is 32-bit. From a 64-bit terminal (including VS Code / Cursor), launch:

```powershell
& "$env:windir\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -File ".\Task-NotesQuery.ps1" -DryRun
```

`-DryRun` scans Notes and writes a preview CSV. It does not create Jira issues, add comments, or update the state file.

### 4. Apply the import

```powershell
& "$env:windir\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -File ".\Task-NotesQuery.ps1"
```

For each mailbox with proposed changes, type `YES` (exactly) to continue, or press Enter to skip that mailbox.

## Import script options

```powershell
.\Task-NotesQuery.ps1 [-DryRun] [-Force] [-MaxIssues <n>] [-ReplacePat]
```

| Parameter | Effect |
| --- | --- |
| `-DryRun` | Scan and write preview CSVs only. No Jira or state changes. |
| `-Force` | Apply Jira changes without the `YES` confirmation. |
| `-MaxIssues <n>` | Abort a mailbox if more than `<n>` new issues would be created. Use this as a safety cap on the first live run. |
| `-ReplacePat` | Prompt for a new Jira PAT and replace the stored credential. |

If you start the import from 64-bit PowerShell, the script exits and prints the 32-bit command to run instead.

## How an import works

For each enabled mailbox the script:

1. Opens the Notes database on `notes.server` using `mailboxes[].databasePath`.
2. Walks `notes.viewName` (for example `By Category`).
3. Keeps documents whose category path matches `includedCategoryPaths`.
4. Skips documents already listed in the mailbox state CSV (`csvPath`).
5. Skips subjects that match `processing.sensitiveSubjectPatterns`.
6. Queues remaining documents in `jiraCategoryPaths` for Jira.
7. Writes a preview CSV, then (unless `-DryRun`) creates issues and comments after confirmation.

Conversation handling:

| Situation | Action |
| --- | --- |
| New email with no processed parent | Create a Jira issue |
| Reply whose parent (or conversation root) already has a Jira key | Add a comment on that issue |
| Reply whose parent was processed earlier but has no Jira key in the state CSV | Skip (`SkipUnmappedParent`) |

New issues get:

- Summary from the Notes subject (with `subjectTextToRemove` stripped)
- Description from `jira.description` plus sender, category, body, and a Notes link
- Project, issue type, labels, and component from the mailbox
- Epic from `mailboxes[].jira.epicKey`, or `jira.defaultEpicKey`
- Due date = Notes created date + `jira.dueAfterDays`
- Start date = Notes created date, if `jira.startDateFieldId` is set

After each successful create or comment, the document is appended to the state CSV so a later run will skip it.

## Configuration

All settings live in `config.json` next to the scripts.

### `modules`

| Key | Description |
| --- | --- |
| `directory` | Folder containing `Security.psm1` and `JiraAPI.psm1`. Usually `Modules`. |

### `security`

| Key | Description |
| --- | --- |
| `credentialTarget` | Windows Credential Manager target name for the Jira PAT. Example: `NotesToJira:JiraPAT`. |

### `jira`

| Key | Description |
| --- | --- |
| `url` | Jira base URL, with no trailing path. Example: `https://jira.example.com` |
| `requestTimeoutSeconds` | HTTP timeout per Jira request. |
| `maximumRetryCount` | Optional. Extra retries for Jira requests. Default is `2` if omitted. |
| `metadataCacheMinutes` | Present in config; reserved for Jira metadata caching. |
| `startDateFieldId` | Optional custom field ID for start date (for example `customfield_13302`). Omit or leave blank to skip. |
| `dueAfterDays` | Days after the Notes created date to set as the Jira due date. |
| `defaultEpicKey` | Fallback Epic issue key (must be an Epic). Used when a mailbox does not set its own `jira.epicKey`. |
| `description.introduction` | Italic intro line in the issue description. |
| `description.action` | Text after **For action:** |
| `description.linkText` | Label for the Notes document link. |

### `notes`

| Key | Description |
| --- | --- |
| `server` | Domino server name as Notes expects it (for example `ServerName/OU/ORG`). |
| `viewName` | View used to walk documents. The first column must be the category path. Typical value: `By Category`. |

### `processing`

| Key | Description |
| --- | --- |
| `sensitiveSubjectPatterns` | Wildcard patterns (`-like` syntax). Matching subjects are skipped and never sent to Jira. |
| `subjectTextToRemove` | Literal strings stripped from the Jira summary (for example classification tags). |
| `emptySubjectText` | Summary used when the subject is empty after stripping. |

### `mailboxes`

Each array entry is one Notes database to scan.

| Key | Description |
| --- | --- |
| `name` | Display name used in console output and preview file names. |
| `enabled` | `true` to process this mailbox; `false` to skip it. If omitted, the mailbox is treated as enabled. |
| `databasePath` | Notes NSF path relative to the server (for example `mail\\shared.nsf`). |
| `csvPath` | State file path. Absolute, UNC, or relative to the application folder. Created if missing. |
| `includedCategoryPaths` | Category paths to read from Notes. |
| `jiraCategoryPaths` | Subset of those paths that should create Jira issues/comments. |
| `includeDescendantCategories` | If `true`, also include categories under the listed paths. If `false` or omitted, only exact path matches. |
| `jira.projectKey` | Jira project key. |
| `jira.issueType` | Issue type name as it appears in Jira (for example `Task`). |
| `jira.labels` | Array of Jira labels. |
| `jira.componentName` | Component name exactly as in Jira (matched case-insensitively). |
| `jira.epicKey` | Optional Epic for this mailbox. Falls back to `jira.defaultEpicKey`. |

Category paths use Notes-style backslashes, for example:

```text
00 Intray\04 General Enquiries\Inbox Name
```

Forward slashes in config are normalised to backslashes.

### Example mailbox

```json
{
  "name": "General Enquiries",
  "enabled": true,
  "databasePath": "mail\\shared-mailbox.nsf",
  "csvPath": "C:\\Data\\NotesToJira\\general-enquiries.csv",
  "includedCategoryPaths": [
    "00 Intray\\04 General Enquiries\\Team Inbox"
  ],
  "jiraCategoryPaths": [
    "00 Intray\\04 General Enquiries\\Team Inbox"
  ],
  "includeDescendantCategories": false,
  "jira": {
    "projectKey": "PROJ",
    "issueType": "Task",
    "labels": ["Query", "Notes_Inbox"],
    "componentName": "Business Tasks",
    "epicKey": "PROJ-1234"
  }
}
```

## Output files

| File | When | Purpose |
| --- | --- | --- |
| `JiraImportPreview-<MailboxName>.csv` | Every scan | Planned actions (`CreateIssue`, `AddComment`, `SkipUnmappedParent`) before changes are applied. Written next to the scripts. |
| State CSV at `mailboxes[].csvPath` | After each successful create or comment | Prevents reprocessing. Do not edit this file unless you intend to change what is considered already imported. |

Preview columns include action, existing issue key, subject, created time, Notes universal IDs, and category path.

State columns include `universalid`, `parent`, `rootUniversalId`, `jiraIssueKey`, `jiraAction`, subject, body, authors, Notes URL, and `processedAt`.

## Security

- The PAT is stored only in Windows Credential Manager for the current user. It is never written to `config.json`, preview CSVs, or console logs.
- Run setup and import as the same Windows account.
- Emails whose subject matches `sensitiveSubjectPatterns` are not imported.
- Keep the state CSV on a path that only the importer account needs to write.

## Troubleshooting

**“This script must run in 32-bit Windows PowerShell”**  
Use the `SysWOW64\WindowsPowerShell\v1.0\powershell.exe` command shown above. Setup can stay 64-bit; only the Notes import cannot.

**Could not open Notes database**  
Confirm Lotus Notes is installed, you can open the NSF in the Notes client, and that `notes.server` and `mailboxes[].databasePath` match that database.

**Notes view was not found**  
Set `notes.viewName` to a view that exists in that NSF. The importer expects the first column to be the category path.

**Component / issue type not found**  
Names must match Jira exactly (except case). Re-run `Setup-JiraEmailImporter.ps1` to list what the PAT user can see on the project.

**Configured issue is not an Epic**  
`jira.defaultEpicKey` and `mailboxes[].jira.epicKey` must be Epic issue keys, not Tasks or Stories.

**Replies skipped (`SkipUnmappedParent`)**  
The parent email is already in the state CSV but has no `jiraIssueKey`. Those replies are not posted as comments. Inspect the preview CSV before a live run.

**Stopped because of `-MaxIssues`**  
The scan proposed more new issues than the cap. Raise `-MaxIssues`, or run `-DryRun` and review the preview first.

**Need a new token**  
```powershell
.\Setup-JiraEmailImporter.ps1 -ReplacePat
```
or
```powershell
.\Task-NotesQuery.ps1 -ReplacePat
```
(still in 32-bit PowerShell for the import script).
