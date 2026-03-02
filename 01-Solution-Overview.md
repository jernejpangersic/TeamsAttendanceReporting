# Teams Attendance & Usage Reporting — Solution Overview

**Date:** March 2, 2026
**Scope:** Proof of Concept (PoC) — validate Graph API attendance extraction and produce Excel reports
**Context:** Middle East crisis response — schools transitioning to fully remote learning
**Customer Profile:** M365 Education tenant (A3 or A5 licensing), ~35,000 teachers, hundreds of thousands of students
**Customer Environment:** School Data Sync (SDS) is already configured. Near-real-time data freshness is required.
**Data Retention:** 30 days (local Excel files; Graph retains reports for 1 year if backfill is ever needed)

---

## Problem Statement

With schools moving to fully remote learning, there is an urgent need to centrally report on Microsoft Teams usage and student attendance with **minimal delay**. The reporting must be available at a **tenant or school level**, rather than requiring each teacher to manually export attendance data.

Teams does not natively support centralised, tenant-level attendance reporting for education scenarios. Existing tools:

- **Attendance feature in meetings:** Teacher-scoped, requires manual export per meeting
- **Teams Admin Center reports:** Aggregate usage only, no per-student per-class detail
- **Education Insights (free, built-in):** Class-level views for individual educators only; no school/district-level roll-ups (see Option 0 note below)

---

## Option 0: Education Insights (Built-in)

> **Note — Insights Premium is no longer available.** The "Education Insights Premium" add-on (formerly included in A5) that provided school-level and district-level roll-up dashboards has been retired by Microsoft. The dedicated documentation pages now return 404. As of March 2026, Microsoft no longer onboards new customers onto this tier.

The **free Education Insights** app in Teams remains available on A1/A3/A5 licenses and provides:

- Per-student digital engagement dashboard (meetings attended, assignments, communication)
- Class-level views for **individual educators** (each educator sees only their own classes)
- Advanced inferences (engagement warnings) when enabled by IT admin

### What It Does NOT Provide

- **No school-level or district-level roll-up views** — there is no way for a school leader or IT admin to see aggregated attendance across all classes/teachers
- **No tenant-wide reporting** — each educator sees only their own class data
- **No exportable data pipeline** — data stays inside the Insights app

### Customer Status

| Requirement | Status |
|---|---|
| School Data Sync (SDS) | **Already configured** ✓ |
| License | A3 or A5 |

### Verdict

Since **SDS is already in place**, educators can use the free Insights app immediately for class-level visibility — worth enabling as a quick win. However, **Insights cannot meet the core requirement** of tenant/school-level attendance reporting with minimal delay. A custom solution (Approach 1) is required.

---

## Approach 1: Granular Class Attendance via Microsoft Graph API ⭐ Recommended

### Overview

Uses the Microsoft Graph API to extract per-meeting attendance data for every teacher in the tenant. Produces a detailed record of which students attended which classes, including join/leave times and duration.

**Data freshness: minutes after each meeting ends.** This is the only approach that meets the customer's minimal-delay requirement while providing school/tenant-level aggregation.

### What It Answers

- Did Student X attend Biology at 10am?
- How long were they present?
- Were they late?
- Which students are absent from specific classes? *(v2 — requires SDS roster comparison)*
- What is the attendance rate per class, per teacher?

### How It Works (PoC)

1. **PowerShell script** runs on-demand or as a scheduled task (locally, or as an Azure Automation runbook later)
2. Loads the cached teacher list from `teachers.json` (synced separately via `Sync-Teachers.ps1`)
3. For each teacher, queries their online meetings for the target date (blind polling)
4. For each meeting, pulls the attendance report with per-participant join/leave data
5. Computes attendance status (Present / Late / Partial); Absent detection deferred to v2 (requires roster data)
6. Writes output to an **Excel file** (.xlsx) using the `ImportExcel` module — one file per run

### API Call Flow

```
Pre:    Load teachers.json (synced separately via Sync-Teachers.ps1)
        → Teacher user IDs

Step 1: GET /users/{teacherId}/onlineMeetings
            ?$filter=startDateTime ge {yesterday} and startDateTime lt {today}
        → Meeting IDs and subjects (skip teacher if empty)

Step 2: GET /users/{teacherId}/onlineMeetings/{meetingId}/attendanceReports
        → Report IDs

Step 3: GET /users/{teacherId}/onlineMeetings/{meetingId}/attendanceReports/{reportId}
            ?$expand=attendanceRecords
        → Per-participant: identity, join/leave times, duration
```

### Required Permissions (Application-Level)

| Permission | Purpose |
|---|---|
| `OnlineMeetingArtifact.Read.All` | Read attendance reports for all users |
| `OnlineMeeting.Read.All` | List meetings for all users |
| `User.Read.All` | Resolve user identities |
| `Group.Read.All` | Read teacher group membership |

All require **admin consent** + an **application access policy** (see Prerequisites).

### Excel Output

| Column | Example |
|---|---|
| Date | 2026-03-01 |
| TeacherName | Ms. Al-Rashid |
| TeacherEmail | a.rashid@school.edu |
| MeetingSubject | Biology - Grade 10 |
| MeetingStart | 2026-03-01 10:00:00 |
| MeetingEnd | 2026-03-01 10:45:00 |
| StudentName | Omar Hassan |
| StudentEmail | o.hassan@school.edu |
| JoinTime | 2026-03-01 10:02:00 |
| LeaveTime | 2026-03-01 10:44:00 |
| DurationMinutes | 42 |
| AttendanceStatus | Present |

Files are saved as `attendance_YYYY-MM-DD.xlsx`. Retain the last 30 days of files; older files can be deleted.

### Attendance Status Logic

| Condition | Status |
|---|---|
| Student not in attendance records | **Absent** *(v2 — requires roster comparison)* |
| Join time > meeting start + 10 minutes | **Late** |
| Duration < 50% of meeting length | **Partial** |
| Otherwise | **Present** |

Thresholds (10 min, 50%) are configurable.

### Scale Considerations

| Metric | Light Day | Typical Day | Peak Day |
|---|---|---|---|
| Total teachers | 35,000 | 35,000 | 35,000 |
| Active teachers (with meetings) | 10,000 (29%) | 20,000 (57%) | 30,000 (86%) |
| Avg meetings per active teacher | 2 | 3 | 4 |
| Total meetings/day | 20,000 | 60,000 | 120,000 |
| API calls (blind polling, 1 app) | ~75,000 | ~155,000 | ~275,000 |
| Detail rows/day | 600,000 | 1,800,000 | 3,600,000 |

> **Note:** Detail rows at peak may exceed Excel's ~1M row limit per sheet. Split output by date range if needed.

### API Throttling & Runtime

- Graph API limit for cloud communications endpoints: ~2,000 requests per 10 minutes per app (~200 req/min sustained)
- Each teacher costs 1 API call to list meetings; each meeting costs 2 additional calls (report list + records)
- Teachers with zero meetings cost exactly 1 call (returns empty, skip to next)

**Single app registration estimates (blind polling all 35K teachers):**

| Scenario | API Calls | Runtime @ 200 req/min |
|---|---|---|
| Light day (10K active, 2 mtg avg) | ~75,000 | ~6.25 hours |
| Typical day (20K active, 3 mtg avg) | ~155,000 | ~12.9 hours |
| Peak day (30K active, 4 mtg avg) | ~275,000 | ~22.9 hours |

**With multiple app registrations (recommended for production):**

| Apps | Light Day | Typical Day | Peak Day |
|---|---|---|---|
| 1 | 6.25 hrs | 12.9 hrs | 22.9 hrs |
| 3 | 2.1 hrs | 4.3 hrs | 7.6 hrs |
| 5 | 1.25 hrs | 2.6 hrs | 4.6 hrs |

- A single app handles light days in an overnight window; typical/peak days require 3–5 app registrations
- Each app registration needs its own client credentials and application access policy
- Partition the teacher list across apps (e.g., 7K teachers per app × 5 apps)
- **Note:** Data is available in Graph within minutes of meeting end. The nightly-batch design is driven by throttling/volume, not data availability.

### Risks & Mitigations

| Risk | Mitigation |
|---|---|
| Graph API throttling at 35K teachers | Multiple app registrations (3–5), exponential backoff, nightly batch window |
| "Meet Now" (ad-hoc) meetings may lack attendance data | Ad-hoc meetings may not appear in `/onlineMeetings` listing; encourage teachers to use **scheduled** or **channel** meetings; document limitation |
| Attendance reports retained for **1 year** in Graph | Comfortable margin for the PoC’s 30-day local retention; backfill is possible if needed |
| Channel meetings return **all** reports for the channel | When listing reports for a channel meeting, Graph returns every meeting's reports in that channel — code must filter by meeting ID / date-time |
| 50-report-per-listing limit | API returns at most 50 most-recent reports per call; busy channels may need pagination or more frequent extraction |
| Organiser can disable attendance report | Teams policy "Attendance and engagement report" defaults to **On, but organisers can turn it off**; consider enforcing via policy if needed |
| Attendees can opt out of being in the report | Teams policy "Include attendees in the report" defaults to **Yes, but attendees can opt out**; admin can override |

---

## Deployment Strategy (PoC)

| Phase | Timeline | Deliverable |
|---|---|---|
| 0 | Day 1 | Enable free Education Insights — class-level visibility for educators (SDS already in place) |
| 1 | Day 1–3 | Azure AD app registration, application access policy, test Graph API auth |
| 2 | Day 2–5 | Build PowerShell scripts: `Sync-Teachers.ps1` + `Get-Attendance.ps1`, write Excel |
| 3 | Day 5–6 | Test with a subset of teachers; validate Excel output |
| 4 | Day 6–7 | Scale to full teacher list; validate at volume |

Total PoC duration: **~7 working days**. If the PoC proves out, a production version can add Azure SQL, SharePoint, Power BI, and alerting later.

---

## Architecture Diagram (PoC)

```
┌───────────────────┐     ┌───────────────────────┐     ┌───────────────────────┐
│  Manual / Scheduled │────▶│  PowerShell Scripts    │────▶│  Microsoft Graph API  │
│  trigger            │     │  (local or Azure Auto) │     │  - Attendance Reports │
└───────────────────┘     └───────────┬───────────┘     └───────────────────────┘
                              │
                    ┌─────────┴────────────┐
                    │                      │
                    ▼                      ▼
          ┌──────────────────┐   ┌──────────────────┐
          │  teachers.json     │   │  Excel File (.xlsx) │
          │  (synced weekly)   │   │  attendance_*.xlsx  │
          └──────────────────┘   └──────────────────┘
```

**Post-PoC production path** (if validated):
- Add multiple app registrations (3–5) to parallelize API calls across 35K teachers
- Consider Call Records webhooks for real-time processing (eliminates blind polling of inactive teachers)
- Add Azure SQL for queryable storage
- Add SharePoint Lists for in-Teams browsing
- Add Power BI dashboards for drill-through analytics
- Add Teams Adaptive Card alerts for daily summaries

---

## Data Output (PoC)

| Data | Output | Format |
|---|---|---|
| Meeting attendance detail | **Excel file** | One row per student per meeting |
| File naming | `attendance_YYYY-MM-DD.xlsx` | One file per extraction run |
| Retention | 30 days | Delete files older than 30 days |

---

## Reporting (PoC)

The Excel file is the primary deliverable. Stakeholders can:
- Open in Excel or Excel Online
- Filter/sort/pivot as needed
- Share via email, Teams chat, or SharePoint document library

**Post-PoC upgrade path:**

| Tier | Tool | When to Add |
|---|---|---|
| Browse | SharePoint List views in Teams | After PoC validates the data |
| Alert | Teams Adaptive Cards | When daily notifications are needed |
| Analyse | Power BI dashboards (Azure SQL) | When drill-through analytics are needed |

---

## Prerequisites Summary (PoC)

| Item | Action |
|---|---|
| Azure AD App Registration | Create in Entra ID portal |
| Application permissions + admin consent | Grant via Entra ID |
| **Application access policy** | Required for `OnlineMeetingArtifact.Read.All` — admin must create a policy and assign it to target users (see [MS docs](https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy)) |
| Teacher security group (or list) | Create/identify group containing all teachers; sync to `teachers.json` via `Sync-Teachers.ps1` |
| PowerShell 7+ | With `Microsoft.Graph` and `ImportExcel` modules installed |
| SDS configuration | **Already configured** ✓ |
| Teams attendance tracking policy | Ensure policy is enabled in Teams admin |

---

## Cost Summary (PoC)

| Component | Monthly Cost |
|---|---|
| PowerShell script (run locally) | $0 |
| Azure Automation (if scheduled) | $0-5 |
| Excel output | $0 (included in M365) |
| **Total incremental cost** | **$0–5/month** |

---

## Key Decision Points for the Customer

| # | Question | Impact |
|---|---|---|
| 1 | A3 or A5 licensing? | Both work for the PoC. A5 includes Power BI Pro for a future production upgrade. |
| 2 | ~~Is SDS configured?~~ | **Resolved: Yes** — SDS is already in place. Free Education Insights can be enabled immediately. |
| 3 | Is there an existing teacher security group? | Yes → immediate use. No → create one before starting. |
| 4 | Is Teams attendance tracking policy enabled? | Must be enabled to capture attendance data. |
| 5 | Has an application access policy been created for the app? | Required for application-permission access to attendance reports. |
| 6 | Where should the Excel files be saved? | Local folder, SharePoint document library, or shared network drive. |
