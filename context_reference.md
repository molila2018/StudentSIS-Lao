# StudentSIS Lao — Project Context Reference

This document is the **single source of truth** for picking up work on this project in any session. Read it cover-to-cover before making non-trivial changes. Update it whenever the structure, rules, or conventions shift — keep it short, scannable, and current.

Last updated: 2026-07-03 (post-autoload)

---

## 1. Project identity

| | |
|---|---|
| **Name** | StudentSIS Lao (`StudentSIS_Lao.exe`) |
| **Path** | `C:\Users\NITRO V15\OneDrive\Desktop\StudentSIS_Lao` |
| **Target school** | ມສ ບຶກທົ່ງ (Buekthong Secondary School, Laos) |
| **Grades supported** | ມ.1 – ມ.4 (Grades 7–10, lower-secondary) — and **only** these |
| **Deployment model** | Offline, single-school, single-machine. SQLite file alongside the EXE. |
| **Language** | Lao throughout the UI; English only for code identifiers and a few admin terms |
| **Database file** | `sis_lao.db` (next to the EXE — created on first launch) |
| **Tech stack** | .NET 8 WPF · `net8.0-windows` · `WinExe` · root namespace `StudentSIS` · assembly `StudentSIS_Lao` |
| **Version** | `1.0.0` (csproj + installer .iss; bump both together) |

---

## 2. Tech stack & dependencies

NuGet packages (see `StudentSIS.csproj`):

| Package | Why |
|---|---|
| `System.Data.SQLite` 1.0.118 | All data access — direct ADO.NET, no ORM |
| `iTextSharp` 5.5.13.3 | PDF rendering (Class Summary fallback path; NU1701 warning is expected) |
| `ClosedXML` 0.102.3 | Excel templates + fresh xlsx generation |
| `DocumentFormat.OpenXml` 2.20.0 | Word template token fill (pinned to 2.x — ClosedXML 0.102.3 was built against 2.x; 3.x changed `IdPartPair` from class to record-struct → runtime failure) |

The test project (`tests/IntegrationTests/`) adds: `System.Text.Encoding.CodePages` (for iTextSharp 5 on .NET 8).

## 3. Architecture & code style

| | |
|---|---|
| **UI** | Imperative C# code-behind (`H.MkGrid`, `H.Btn`, etc. in `Views/RemainingPages.cs`). A few pages use XAML (`ReportPage`, `ScoresPage`, `StudentsPage`, `LoginWindow`, `StudentFormWindow`). |
| **Data layer** | Single `static class DB` in `Data/DB.cs`. All SQL lives here or in pages (no ORM, no Repository pattern). |
| **Business logic** | Transaction Script. Methods do a sequence of SQL operations + return a DataTable / scalar. |
| **MVVM** | **No.** Event-driven code-behind; `btn.Click += ...`. Some `INotifyPropertyChanged` (e.g. `MonthlyRow`) but no ViewModels. |
| **DI** | None — static `DB` everywhere. |
| **Formatting** | Compact / one-liner methods are common in `RemainingPages.cs`; the newer code (ReportPage, UserFormWin) is more spacious. |
| **Comments** | Default to no comments. Add only when the *why* is non-obvious (hidden constraint, workaround, surprising behaviour). |

See user feedback memory: project owner is a beginner JS dev learning by reading the code; prefer clarity over cleverness when adding new code. They communicate in Thai but the UI is Lao.

---

## 4. Sidebar pages (14 entries, indices 0–13)

Defined in `Views/MainWindow.xaml.cs` (`_pages` array):

| Idx | Page | Class | Admin only |
|---|---|---|---|
| 0 | ໜ້າຫຼັກ (Dashboard) | `DashboardPage` | no |
| 1 | ໜ້າຫ້ອງຮຽນ (ClassHub) | `ClassHubPage` | no |
| 2 | ຂໍ້ມູນນັກຮຽນ (Students) | `StudentsPage` | no |
| 3 | ລົງທະບຽນ (ລາຍຄົນ) — per-student enrolment | `EnrollmentPage` | no |
| 4 | ລົງທະບຽນ (Batch) — class-wide enrolment | `BatchEnrollPage` | no |
| 5 | ບັນທຶກຄະແນນພາກຮຽນ (semester / final exam scores) | `ScoresPage` | no |
| 6 | ບັນທຶກຄະແນນປະຈຳເດືອນ (monthly — academic + per-month CHA1/LAB1) | `MonthlyScoresPage` | no |
| 7 | ໃບຄະແນນ / ລາຍງານ (Reports — non-score only: Enrollment Agreement + Profile) | `ReportPage` | no |
| 8 | ຂຶ້ນຊັ້ນ / ຈົບ (Promotion) | `PromotionPage` | yes |
| 9 | ວິຊາ (Subjects) | `SubjectsPage` | no |
| 10 | ຈັດການຜູ້ໃຊ້ (Users) | `UsersPage` | yes |
| 11 | ປີການສຶກສາ (AcademicYear) | `AcademicYearPage` | yes |
| 12 | ຕັ້ງຄ່າລະບົບ (Settings) | `SettingsPage` | yes |
| 13 | ປະຫວັດຄະແນນນັກຮຽນ (Score History) | `ScoreHistoryPage` | no |

Non-admin (teacher) users have entries 8, 10, 11, 12 hidden via `Visibility=Collapsed` (see `MainWindow.xaml.cs` ctor). Entry 13 is visible to teachers — it is read-only.

---

## 5. Reports (2 types in `ReportPage`)

All score-related reports were migrated to **`ScoreHistoryPage`** (idx 13). `ReportPage` now keeps only the two non-score administrative reports. **Current dispatch (`Views/ReportPage.xaml.cs`):**

| Idx | Report | Builder | Output |
|---|---|---|---|
| 0 | 📑 ໃບສັນຍາເຂົ້າຮຽນ (Enrollment Agreement) | `GenEnrollmentAgreement` → `Templates/ໃບສັນຍາ.docx` | docx · Word COM/LibreOffice → PDF |
| 1 | 📋 ລາຍງານປະຫວັດນັກຮຽນ (Profile) | `GenStudentProfileReport` → `Templates/ລາຍງານປະຫວັດນັກຮຽນ.docx` | docx · Word/LibreOffice → PDF |

**Score reports** (Monthly · Sem 1 · Sem 2 · Annual · Individual) live exclusively on the Score History page via these `internal static` entry points in `ReportPage.xaml.cs` — see §10 for full UI:

| Entry point | Template | Used by |
|---|---|---|
| `RenderClassMonthlyXlsx` | `ໃບຄະແນນ.xlsx` Sheet 1 | `ClassHistoryWindow` 📚 |
| `RenderClassSemesterXlsx(sem)` | `ໃບຄະແນນ.xlsx` Sheet 2/3 | `ClassHistoryWindow` 📚 |
| `RenderClassAnnualXlsx` | `ໃບຄະແນນ.xlsx` Sheet 4 | `ClassHistoryWindow` 📚 |
| `RenderIndividualMonthlyXlsx` | `ລີພອດຄະແນນບຸກຄົນ.xlsx` | `StudentHistoryWindow` 📖 |
| `RenderIndividualSemesterXlsx(sem)` | `ລີພອດຄະແນນບຸກຄົນ.xlsx` | `StudentHistoryWindow` 📖 |
| `RenderIndividualAnnualXlsx` | `ລີພອດຄະແນນບຸກຄົນ.xlsx` | `StudentHistoryWindow` 📖 |

**PDF strategy:** for Excel-based reports the xlsx is generated then converted via `ConvertXlsxToPdfViaExcel` (Excel COM `ExportAsFixedFormat` — pixel-identical to Excel print). For docx reports we use Word COM `SaveAs2(wdFormatPDF=17)` or LibreOffice `soffice.exe --headless --convert-to pdf` as fallback.

### Template files (`Templates/`)

| File | Used by | Notes |
|---|---|---|
| `ໃບຄະແນນ.xlsx` | All class-history score exports (Sheet 1/2/3/4) | 4 sheets; layout: row 1 national name · row 2 patriotic motto · row 3 school name · row 4 dynamic report title (`ສະຫຼຸບ...`) · row 5 headers · **row 6+ data with formulas** · **row 10 signatures** (`ອຳນວຍການ` / `ວິຊາການ` / `ຄູປະຈຳຫ້ອງ`, merged across 3 column spans). `FillSheet` finds the header row dynamically by scanning for `ລຳດັບ`, snapshots the signature row's cells + merge spans before stripping data rows, then re-emits the signature row two rows below the last student. Rank column is overwritten with a computed value by `WriteClassRanks` post-pass. |
| `ໃບສັນຍາ.docx` | Enrollment Agreement (ReportPage idx 0) | `{{TOKEN}}` placeholders filled via `FillDocxTokens` |
| `ລີພອດຄະແນນບຸກຄົນ.xlsx` | All individual-history score exports | v2 layout — score column at **F**, subjects rows 8–19, summary 20–22, CHA/LAB 23/24, **4 signatures at row 27** (A=ອຳນວຍການ · C=ຜູປົກຄອງ · E=ວິຊາການ · G=ຄູປະຈຳຫ້ອງ). Constants in `TemplateSubjectRows` + `IndScoreCol/IndSumRow/...` |
| `ລາຍງານປະຫວັດນັກຮຽນ.docx` | Profile (ReportPage idx 1) | Patched: original had `{{MOTHER_PROVIN`+`CE` (no closing braces) — repacked via `tests/IntegrationTests/Repack.cs` |

All four templates ship next to the EXE via `<None Update="Templates\..."> <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory></None>` in csproj.

### Shared helpers in `ReportPage.xaml.cs`

- `ResolveLaoFontName()` — runtime font detection via `Fonts.SystemFontFamilies`. Preference order: Phetsarath OT → Saysettha OT → Noto Sans Lao → Lao UI → Leelawadee UI (Windows fallback).
- `ContainsLao(string)` — true if any char is in U+0E80–U+0EFF.
- `FillDocxTokens(string, Dictionary<string,string>)` — paragraph-level merge-and-replace. **Robust to Word splitting tokens across runs.** Used by both `GenEnrollmentAgreement` and `GenStudentProfileReport`.
- `ExportTablePdf(...)`, `ConvertXlsxToPdfViaExcel(...)`, `ConvertDocxToPdf(...)` — output helpers.
- `FillSheet` / `MapSheet` / `ProduceTemplateReport` / `WriteClassRanks` / `BuildIndividualXlsx` / `ComputeIndividualScores` — score-template fill machinery, called from the `Render*Xlsx` entry points.

---

## 6. Database schema

Defined inline as a multi-line SQL string in `Data/DB.cs` `Initialize()` (~line 47). Migrations follow as separate `ALTER`/`CREATE` blocks (idempotent, wrapped in try/catch).

| Table | Purpose | Key columns |
|---|---|---|
| `Users` | Logins | `UserID`, `Username UNIQUE`, `Password` (plaintext — single-school offline OK), `FullName`, `Role` (`admin`/`teacher`), `IsActive`, `LastLogin` |
| `Students` | 24+ identity fields | `StudentID`, `StudentCode UNIQUE`, `FirstName`, `LastName`, `Gender`, `BirthDate`, birth place (3), current address (3), father (7), mother (7), legacy `ParentName`/`ParentPhone`, `GradeLevel NOT NULL`, `ClassRoom`, `AcademicYear NOT NULL`, `Status` (default `ກຳລັງຮຽນ`) |
| `Subjects` | 14 official MoES subjects | `SubjectID`, `SubjectCode UNIQUE` (CIV1, SCI1, LAO1, MATH1, ENG1, ICT1, GEO1, HIS1, ART1, MUS1, PE1, VOC1, **CHA1**, **LAB1**), `SubjectName`, `GradeLevel` ("" = grade-agnostic), `SortOrder` |
| `Enrollments` | Year-wide registration | `EnrollID`, `StudentID FK`, `SubjectID FK`, `AcademicYear`, `Semester` (1/2), `Teacher`, `UNIQUE(StudentID,SubjectID,AcademicYear,Semester)` |
| `Scores` | Semester final scores | `ScoreID`, `EnrollID UNIQUE FK`, `MidScore`, `FinalScore`, `TotalScore`, `Level`, `Remarks`, `UpdatedAt` |
| `MonthlyAssessments` | Per-month entries | `MonthlyID`, `EnrollID FK`, `Month`, `ActivityScore` (/3), `DisciplineScore` (/2), `HomeworkScore` (/5), `UNIQUE(EnrollID,Month)` |
| **`EvaluationScores`** | Manual CHA1/LAB1 per context | `EvalID`, `StudentID FK`, `AcademicYear`, `Context` (`Month1`..`Month8` / `SEM1` / `SEM2` / `ANNUAL` — `Month{N}` via `DB.MonthContextName`), `SubjectCode`, `Score`, `UNIQUE(StudentID,AcademicYear,Context,SubjectCode)` |
| `AttendanceRecords` | Per-day attendance | (feature removed from UI but table kept) |
| `GradeHistory` | Promotion log | `StudentID`, `FromGrade`, `ToGrade`, `AcademicYear` (= year being promoted INTO), `ClassRoom` (= room while in `FromGrade`), `Note`, `ChangedBy`, `ChangedAt` |
| `AcademicYears` | Year registry | `Year PK`, `IsCurrent`, `StartDate`, `EndDate`, `Note`, `CreatedBy`, `CreatedAt` |
| `ActivityLog` | Audit trail | `UserID`, `Username`, `Action`, `Detail`, `LoggedAt` |
| `Settings` | Key/value config | `school_name`, `current_year`, `current_semester`, `mid_pct`, `final_pct`, `pass_score` |
| `Announcements`, `Conversations`, `Messages` | Mostly unused scaffolding | — |

Migrations (also in `Initialize()`):
1. Adds Students columns for Birth*, Father*, Mother* (idempotent ALTERs).
2. Backfills `FatherName`/`FatherPhone` from legacy `ParentName`/`ParentPhone`.
3. Migrates `mid_pct=60` → 50 and `final_pct=40` → 50 (one-time).
4. Backfills `AcademicYears` from distinct years in Students/Enrollments/GradeHistory.
5. Creates `EvaluationScores` table (newer migration).
6. Adds `GradeHistory.ClassRoom` column + backfills it from `Students.ClassRoom` in two passes:
   1. Pass 1 — exact: rows where `AcademicYear` matches the student's current `AcademicYear`.
   2. Pass 2 — best-effort: any remaining NULL/empty rows, using the student's current room as a guess. This is critical for GRADUATED students (whose final GradeHistory row has `AcademicYear = year-after-grad` and never matches their `Students.AcademicYear = year-finished`). Without pass 2, graduates would be invisible in the score-history roster.

---

## 7. Business rules — score computation

**The 14 subjects** (per `Subjects` table seed) split into:
- **12 academic subjects** included in totals/averages/rankings: LAO1, CIV1, MATH1, SCI1, GEO1, HIS1, ENG1, ICT1, PE1, MUS1, ART1, VOC1
- **2 evaluation-only**: **CHA1** (ຄຸນສົມບັດ / character) and **LAB1** (ການອອກແຮງງານ / labor). **NEVER auto-computed, averaged, or summed** — manually entered.

**Helpers in `DB.cs`** (memorise these; they're called everywhere):

| Method | Returns | Rule |
|---|---|---|
| `CalcMonthlyTotal(a,d,h)` | double | `a + d + h` (sum of activity/3 + discipline/2 + homework/5 = /10) |
| `CalcTotal(mid, fin)` | double | `mid×(MidPct/100) + fin×(FinalPct/100)` — default 50/50 |
| `CalcLevel(total)` | string | 4-band: ≥8 ດີຫຼາຍ · ≥6 ດີ · ≥PassScore ຜ່ານ · else ບໍ່ຜ່ານ |
| `CalcMoESLevel(total)` | string | 5-band: ≥9 ດີຫຼາຍ · ≥7 ດີ · ≥5 ປານກາງ · ≥3 ອ່ອນ · else ຕົກ |
| `SemesterForMonth(m)` | 1/2 | Sept–Jan → 1, Feb–June → 2 |
| `MonthsInSemester(s)` | int[] | sem 1 → {9,10,11,12}, sem 2 → {2,3,4,5} |
| `FinalExamMonth(s)` | int | sem 1 → 1 (Jan), sem 2 → 6 (June) |
| `NextGrade(g)` | string | ມ.1→ມ.2 · ມ.2→ມ.3 · ມ.3→ມ.4 · ມ.4→ຈົບ |
| `RecomputeMidFromMonthly(enrollId)` | void | Averages MonthlyAssessments → updates `Scores.MidScore` + recomputes `TotalScore` + `Level`. **Early-returns for CHA1/LAB1** — never auto-computes them. |
| `GetEvaluationScore(sid, year, ctx, code)` | double? | Pull manual CHA/LAB score |
| `SetEvaluationScore(...)` | void | Upsert; null score deletes |
| `GetClassRanking(...)` | DataTable | `ORDER BY ສະເລ່ຍ DESC, ລວມ DESC`. Aggregates EXCLUDE CHA1/LAB1. **Does NOT push failed students last** — the "failed → ‘ຕົກ’ rank" rule is layered on top in report builders. |
| `GetStudentHistoryYears(sid)` | DataTable | Every (year, grade-at-the-time, room-at-the-time) the student has data for. Reconstructs historical grade/room from GradeHistory (preferred — smallest `AcademicYear > Y` describes year Y via `FromGrade`+`ClassRoom`) → current `Students` row (fallback for the most recent year). |
| `GetHistoryMonthly(sid, year)` | DataTable | Per (subject × month) rows for that year. Academic subjects come from `MonthlyAssessments`; CHA1/LAB1 come from `EvaluationScores(Month1..Month8)` via a UNION (they don't live in `MonthlyAssessments` so an INNER JOIN silently dropped them — the UNION restores them with NULL sub-scores and the eval score in `ລວມເດືອນ`). CHA1/LAB1 rows are sorted LAST within each month. |
| `GetHistorySemester(sid, year, sem)` | DataTable | Per-subject mid/final/total/level for that (year, sem). CHA1/LAB1 rows show the manually-entered `EvaluationScores.Score` in the total column instead of derived numbers. |
| `GetHistorySemesterSummary(sid, year, sem)` | (subjects, avg, total, rank, classSize, failed) | One-row summary with class rank computed against historical classmates (cohort = students whose own GradeHistory or current Students row puts them in the same grade for that year). EXCLUDES CHA1/LAB1. Failed students get `failed=true` and the report layer renders the rank as ‘ຕົກ’. |
| `GetHistoryAnnualSummary(sid, year)` | (sem1Avg, sem2Avg, annualAvg, rank, classSize, level, failed) | Annual readout. `annualAvg` = mean of the two sem averages (falls back to whichever has data if one is empty). `rank` ranks the student against the historical cohort by that same annual mean. `level` = MoES 5-band for `annualAvg`. CHA1/LAB1 excluded throughout. |
| `GetHistoricalClassRoster(year, grade, room, statusFilter=null)` | DataTable | Every student who was in (year, grade, room) historically. UNION of (a) GradeHistory rows whose smallest `AcademicYear > year` row has `FromGrade=grade` + `ClassRoom=room` and (b) Students still active in that year with matching current grade+room AND no later GradeHistory. Handles graduated + promoted + current students. Columns: `StudentID`, `ລະຫັດ`, `ຊື່ນັກຮຽນ`, `ເພດ`, `ສະຖານະ`. **`statusFilter`**: optional `Students.Status` match (`"ກຳລັງຮຽນ"` / `"ຈົບ"` / `"ອອກ"`); `null` returns all statuses. Used by the score-history page's Active / Graduated / All radios. |
| `GetClassMonthGrid(year, grade, room, month)` | DataTable | Class-wide per-month grid for `ClassHistoryWindow`. Rows = roster students; columns = academic subject codes; cells = that subject's monthly /10 total. CHA1/LAB1 columns deliberately omitted. |
| `GetClassSemesterGrid(year, grade, room, sem)` | DataTable | Same shape as `GetClassMonthGrid` but cells = `Scores.TotalScore` for that (student, subject, sem). Used by `ClassHistoryWindow` for S1/S2 views. CHA1/LAB1 columns omitted. |
| `GetClassAnnualGrid(year, grade, room)` | DataTable | Same shape; cells = mean of `Scores.TotalScore` across sem1+sem2. Used for the annual view. CHA1/LAB1 columns omitted. |
| `GetClassSemesterSummary(year, grade, room, sem)` | DataTable | One row per roster student with `ສະເລ່ຍ / ລວມ / ອັນດັບ / ລະດັບ` for that semester. CHA1/LAB1 excluded. Still called by `RenderClassSemesterXlsx` (export path). |
| `GetClassAnnualSummary(year, grade, room)` | DataTable | One row per roster student with `ສະເລ່ຍພາກ 1/2/ປະຈຳປີ / ອັນດັບ / ລະດັບ`. CHA1/LAB1 excluded. Still used by the export path. |
| `GetHistoryAnnual(sid, year)` | DataTable | Per-subject annual scorecard for one student: `ລະຫັດວິຊາ · ຊື່ວິຊາ · ສະເລ່ຍພາກ1 · ສະເລ່ຍພາກ2 · ສະເລ່ຍປະຈຳປີ`. Academic = mean of sem1+sem2 `TotalScore`. CHA1/LAB1 pulls per-context values from `EvaluationScores(SEM1/SEM2/ANNUAL)`. Used by `StudentHistoryWindow`'s annual view. |
| `Backup(dest)` / `Restore(src)` | void | Plain `File.Copy` against `sis_lao.db` |
| `SetCurrentAcademicYear(y)` | void | Canonical year-switch: updates registry IsCurrent + Settings.current_year + resets semester to 1 + reloads cache. Use this — never write to `current_year` directly. |

**Pass/fail rule (CHA1/LAB1 contract):**
- `RecomputeMidFromMonthly` early-returns for CHA1/LAB1 — verified by integration test.
- Every score-aggregate query has `AND sub.SubjectCode NOT IN ('CHA1','LAB1')`.
- Failed students' rank: in the report-layer (not in `GetClassRanking`), students with any subject below `PassScore` get rank `"ຕົກ"` (text). Passing students get the tie-aware numeric rank (1224 style).

**Academic calendar:**
- Sem 1: monthly entries Sept · Oct · Nov · Dec → final exam January
- Sem 2: monthly entries Feb · Mar · Apr · May → final exam June
- Enrolment is year-wide — every subject registration creates 2 rows (Sem 1 + Sem 2).

---

## 8. Score-entry pages — semantics

Both score pages now use the **same class-roster × one-subject** workflow: teachers pick (Year, Grade, Room, Subject, Month-or-Sem), the grid loads the whole class for that one subject, scores get filled down the column with Enter, and Ctrl+S saves everything in one transaction. Switch the Subject combo to start the next subject.

| Page | What it edits | CHA/LAB display | Read-only mode |
|---|---|---|---|
| `MonthlyScoresPage` (idx 6) | **Academic** → `MonthlyAssessments` (Activity/Discipline/Homework split for the picked month). **CHA1/LAB1** → `EvaluationScores(Month{N}, code)` single /10 score. | CHA1/LAB1 appear in the Subject combo just like academic subjects; picking one swaps the grid to a single /10 `_colEval` column | Status ≠ ‘ກຳລັງຮຽນ’ → grid `IsReadOnly=true` + Save disabled |
| `ScoresPage` (idx 5) | **Academic** → `Scores.FinalScore` UPSERT (Mid preserved — owned by monthly recompute). Total + Level recomputed from Mid + new Final. **CHA1/LAB1** → `EvaluationScores(SEM{N}, code)` UPSERT for ພາກ 1 / ພາກ 2, or `EvaluationScores(ANNUAL, code)` UPSERT for ປະຈຳປີ. | CHA1/LAB1 in the Subject combo; picking one swaps the grid from `Mid · Final · Total · Level` to a single `ຄະແນນ (/10)` column (header becomes `ຄະແນນປະຈຳປີ (/10)` when ພາກ = ປະຈຳປີ). Selecting ພາກ = ປະຈຳປີ auto-filters the Subject combo to CHA1/LAB1 only (Import/Download buttons disabled since the SIS_LAO_SEMESTER_v4 template covers ພາກ 1/2 only). | Same read-only enforcement |

**Auto-load (both pages):** the `🔄 ໂຫຼດ` button was removed; the roster loads automatically on page open and on every filter change. `_autoReload` initialised to `true`. F5 still triggers a manual reload for the keyboard-first user. Initial-load path: if `DB.NavGrade` is set (arrived via ClassHub), `ApplyNavContext` fires; otherwise the constructor kicks a `LoadRoster(manual: false)` via `Dispatcher.InvokeAsync` so the default filter tuple lands the class grid immediately.

**Keyboard-first input (both pages):** Enter on the score cell **commits + moves selection DOWN** to the same column on the next row — the primary grading flow. Ctrl+S saves. F5 reloads the roster. Type/Tab/Enter cycle lets a teacher grade 30 students in 30 keystrokes per subject.

**Per-student view (the old `ScoresPage` layout) was removed** — that workflow lives on the Score History page (sidebar idx 13): pick a student, press 📖, see the multi-year per-subject breakdown.

**CHA1 / LAB1 contract:** Subject combo behavior is identical to any other subject — pick CHA1 or LAB1 and the grid swaps to a single /10 score cell. All values land in `EvaluationScores` with the appropriate context: `Month1`..`Month8` for per-month entries, `SEM1` / `SEM2` for semester, `ANNUAL` for year-end. Each context is independent (no derivation across periods). `DB.SetEvaluationScore` UPSERT — `score=null` deletes the row. Reports + class history reads CHA1/LAB1 monthly from `EvaluationScores(Month{N})` (never sums `MonthlyAssessments`); semester reads use `SEM1`/`SEM2`; annual reads use `ANNUAL`. CHA1/LAB1 are never included in averages, rankings, semester totals, or annual totals (all academic queries `NOT IN ('CHA1','LAB1')`).

**Score input** (both pages): direct keyboard entry via `DataGridTextColumn`, **integers only** (no decimals). Decimal separators (`.` / `,`), letters, and signs are blocked at keystroke time by a `PreviewTextInput` filter attached via `PreparingCellForEdit`. On commit, `DB.TryParseIntScore(text, 0, max, out value)` enforces strict integer parsing — invalid input (decimals, out-of-range, non-numeric paste) cancels the commit and calls `CancelEdit` so the cell **restores the previous value** rather than holding the bad text. Per-column max: Activity=3 / Discipline=2 / Homework=5 / Eval=10 / Final=10. Empty input maps to 0. (`DB.TryParseScore` decimal-capable helper still exists for any future caller that needs it.)

**EvaluationScores entry flow on MonthlyScoresPage:** the page exposes Grade · Room · Status · Subject · Year · Month (no ປະເພດ/Type dropdown). Picking CHA1 or LAB1 swaps the grid to a single /10 column and reads/writes `EvaluationScores(Month{N}, code)` via `DB.GetEvaluationScore` / `SetEvaluationScore`. SEM1 / SEM2 / ANNUAL CHA1/LAB1 values are entered on `ScoresPage` (the `ພາກ` combo has ພາກ 1 · ພາກ 2 · ປະຈຳປີ; ປະຈຳປີ writes context `ANNUAL`).

**Excel score import (both pages — PER-SUBJECT):** in addition to manual entry, teachers can bulk-import scores via two buttons in each page's filter bar — `📥 ດາວໂຫຼດແບບຟອມ` (Download Template) and `📤 ນຳເຂົ້າ Excel` (Import Excel). Each download produces a template containing **one subject only**; each upload validates against that same subject. Both flows live in [`Data/ExcelImport.cs`](Data/ExcelImport.cs) and share `ImportPreviewWindow` (in `RemainingPages.cs`).

| | Monthly + academic | Monthly + CHA1/LAB1 | Semester (any subject) |
|---|---|---|---|
| Scope key | Year · Grade · Room · Month · **SubjectCode** | same | Year · Grade · Room · Semester · **SubjectCode** |
| Subject picker | MonthlyScoresPage's `_subject` combo | same | ScoresPage's `CmbImportSubject` (only feeds the import buttons; per-student view unchanged) |
| Template columns | ລຳດັບ · ລະຫັດ · ຊື່ · **ກິດຈະກຳ2 · ຮ່ວມຮຽນ3 · ກວດກາ5** (3 sub-scores, per-column max 2/3/5) | ລຳດັບ · ລະຫັດ · ຊື່ · ຄະແນນ (/10) | ລຳດັບ · ລະຫັດ · ຊື່ · **ຄະແນນສອບເສງພາກຮຽນ (/10)** (Final only — no Midterm) |
| Roster rows | one per active student in (year, grade, room) | same | same |
| Save | `MonthlyAssessments` UPSERT — Discipline/Activity/Homework written **verbatim** (no split; columns map 1:1 to DB) | `EvaluationScores(Month{N}, code)` UPSERT | Academic → `Scores` UPSERT writing only `FinalScore` (`MidScore` preserved — owned by `RecomputeMidFromMonthly`). Then Total + Level recomputed from existing Mid + new Final. CHA1/LAB1 → `EvaluationScores(SEM{N}, code)` UPSERT |
| Post-save | `RecomputeMidFromMonthly` fires for every academic enrollment touched | none | none |

**Visual styling (matches the official Lao monthly-score sheet):**
- Font: **Phetsarath OT** throughout (title 14pt bold, headers 11pt bold, data 11pt)
- Title: row 1 column A, no merge (overflows into adjacent empty cells)
- Headers: row 3, light gray fill (`#D3D3D3`), bold, centered, thin border
- Data cells: row 4+, thin border, centered for ordinal/code/score columns, left-aligned for name
- Default row height: 18pt · Title row: 23.4pt
- Page setup: Portrait, A4, fit-to-1-page-wide, margins L/R/T/B = 0.75/0.75/0.75/0.5, centered horizontally; rows 1–3 repeat on print (so multi-page rosters keep their header)
- Column widths: A=6 · B=13 · C=30 · monthly academic D/E/F=12 · monthly eval D=13 · semester D/E=22 · Z (hidden) = 24

**Template format:** row 1 = title (includes `ວິຊາ <CODE> — <Name>`), row 2 blank, row 3 = headers, row 4+ = data. Hidden metadata column carries the round-trip key:
- Monthly templates → column **Z** (6 visible cols, A-F)
  - `Z1` = `SIS_LAO_MONTHLY_v3` · `Z2`-`Z6` = year / grade / room / month / subject
- Semester templates → column **Y** (4 visible cols, A-D)
  - `Y1` = `SIS_LAO_SEMESTER_v4` · `Y2`-`Y6` = year / grade / room / sem / subject

Parser rejects:
- Workbook whose magic doesn't match (random external xlsx files cannot be imported)
- Any (year, grade, room, month/sem) scope mismatch
- **Subject-scope mismatch** — a template generated for MATH1 cannot be uploaded under ENG1

**Validation rules** (Lao status messages shown in the preview):
- Blank score rows are silently skipped (not invalid, not saved) — protects existing scores when a teacher fills only a few cells
- `❌ ບໍ່ພົບນັກຮຽນ` — StudentCode not in the (year, grade, room) roster
- `❌ ນັກຮຽນ ... ມີສະຖານະ ‘ຈົບ/ອອກ’` — graduates and withdrawn students are read-only (status must be `ກຳລັງຮຽນ`)
- `❌ ບໍ່ໄດ້ລົງທະບຽນວິຊານີ້` — no `Enrollments` row for (StudentID, SubjectID, Year, Sem); CHA1/LAB1 skip this check
- `❌ ກິດຈະກຳ/ຮ່ວມຮຽນ/ກວດກາຜິດ` — sub-score fails `DB.TryParseIntScore(0, max)` where max is 2/3/5 respectively (monthly academic)
- `❌ ຄະແນນຜິດ` — single score fails `DB.TryParseIntScore(0, 10)` (monthly CHA1/LAB1, semester)
- `❌ ຊ້ຳກັນ` — same StudentCode appears more than once in the file (keep first, reject rest)
- Subject identity is enforced by the file-level magic/`Z6` check; no per-row "subject mismatch" because the data rows no longer carry a Subject column

**Preview window** (`ImportPreviewWindow`): DataGrid columns adapt to the (Kind, IsEvalSubject) combo on the result:
- Monthly + academic → `ລຳດັບ · ລະຫັດ · ຊື່ · ກິດຈະກຳ (/2) · ຮ່ວມຮຽນ (/3) · ກວດກາ (/5) · ລວມ (/10) · ສະຖານະ`
- Monthly + CHA1/LAB1 → `ລຳດັບ · ລະຫັດ · ຊື່ · ຄະແນນ (/10) · ສະຖານະ`
- Semester (any subject) → `ລຳດັບ · ລະຫັດ · ຊື່ · ສອບເສງພາກຮຽນ (/10) · ສະຖານະ`

Header title states the subject (`ວິຊາ MATH1 — ຄະນິດສາດ`). Summary footer counts `✅ ຖືກຕ້ອງ` vs `❌ ຜິດພາດ`. Invalid rows tinted red (`#FEF2F2`). `💾 ບັນທຶກ N ແຖວ` button disabled when ValidCount == 0; on confirm, `ExcelImport.SaveImport` runs every valid row in one transaction. After save, the caller pages refresh only the affected subject's view (no full page rebuild).

**CHA1/LAB1 contract held throughout the import path:** academic monthly saves trigger `RecomputeMidFromMonthly`; CHA1/LAB1 monthly saves do not (the helper itself early-returns for those codes). CHA1/LAB1 semester rows never produce a `Scores` row — only `EvaluationScores`. Verified by 58 integration assertions including standalone MATH1 (3 sub-scores written verbatim) / CHA1 / LAB1 / ENG1 imports, per-column max enforcement, and visual verification against the official Lao monthly-score sheet via `dotnet run --project tests/IntegrationTests -- gen-tpl <out.xlsx>`.

---

## 9. Enrolment pages — current design

Both enrolment pages are now **one-click** (no subject picker — the user picks a student or class and every subject for that grade is enrolled automatically):

| Page | Idx | Flow |
|---|---|---|
| `EnrollmentPage` (per-student) | 3 | year + grade + student → **⚡ ລົງທະບຽນທຸກວິຊາ** → confirm → for each subject in `Subjects` matching the grade (or grade-agnostic), INSERT OR IGNORE 2 rows (sem 1 + sem 2). Reports new/existed counts. |
| `BatchEnrollPage` (class-wide) | 4 | year + grade + room → see the roster → **⚡ ລົງທະບຽນທຸກນັກຮຽນ** → confirm → same loop × every active student in the filter. Status filter blocks graduates. |

The old subject-picker modal (`EnrollPickWindow`/`PickItem`) and the two-checkbox-list UI were deleted. `INSERT OR IGNORE` against `UNIQUE(StudentID,SubjectID,AcademicYear,Semester)` makes the operation idempotent — safe to re-click.

---

## 10. Major user-visible workflows

### Promotion (`PromotionPage`, idx 8, admin)

Redesigned UI: filter card (Year + Grade + Room + status radios) → action toolbar → roster DataGrid (checkbox + 7 columns) → history panel.

**Four bulk actions:**

- **⬆️ Promote Selected** — opens `PromotionConfirmWindow` (destination dropdowns default to `NextYear` / `NextGrade` / same room; per-student preview `Code | Name | From → To`; editable before confirm). Writes `Students.GradeLevel/ClassRoom/AcademicYear` + `GradeHistory` row with `ClassRoom` and `Note='ຂຶ້ນຊັ້ນ'`. **Guardrail:** `EnsureNextYearExists` runs first — if `NextYearString(srcYear)` isn't in the `AcademicYears` registry, refuse to open the dialog; offer to launch `AcademicYearFormWin` prefilled for that year so the admin creates it inline.
- **🔁 Repeat Selected (ຊ້ຳຊັ້ນ)** — same confirm dialog as Promote, but `dstGrade` defaults to the **same** grade (student retained). `AcademicYear` advances to `NextYearString(srcYear)`; `ClassRoom` editable if the school reshuffles sections. `Status` stays `'ກຳລັງຮຽນ'`. `GradeHistory` row: `FromGrade=ToGrade=srcGrade`, `Note='ຊ້ຳຊັ້ນ'`. The dst-grade dropdown hides ‘ຈົບ’ in repeat mode (a retained student can't also graduate). Same active-only filter — graduates + withdrawn skipped. Same `EnsureNextYearExists` guardrail as Promote. Log tag `Repeat` vs `Promotion`.

**`PromotionConfirmWindow`:** year dropdown lists ONLY registered `AcademicYears`. The old auto-add-if-absent fallback was removed — the pre-check above guarantees the target year exists before the dialog opens.

**`AcademicYearFormWin(string? prefillYear = null)`:** the constructor accepts an optional prefill so the promotion guardrail can open it pre-filled with the exact needed year; the plain "➕ ເພີ່ມປີໃໝ່" button on `AcademicYearPage` still passes null (defaults to `NextYearString(CurrentYear)`).
- **🎓 Graduate Selected** — simple Yes/No confirm. Sets `Students.Status='ຈົບ'`; `GradeLevel` stays anchored to finishing grade, `AcademicYear` stays at finishing year. `GradeHistory` row recorded with `ToGrade='ຈົບ'` and `AcademicYear = year-after-grad`.
- **🔄 Promote Entire Classroom** — auto-selects every `Status='ກຳລັງຮຽນ'` row in the current cohort, then funnels through the same Promote-Selected confirm flow. Already-graduated students are skipped.

Plus convenience buttons: `📋 Select All` (active only) / `❌ Clear Selection`.

**Performance:** roster query (`DB.GetHistoricalClassRoster` + enrich) fires only when Year + Grade + Room are all picked.

**Data preservation rules:**
- **No data is deleted** — `Enrollments`, `Scores`, `MonthlyAssessments`, `EvaluationScores`, `AttendanceRecords` all preserved.
- CHA1/LAB1: never recalculated, never modified.
- Graduates' historical data is reachable via the score-history page by switching the status radio to `🎓 ຈົບ`.

**Status radios** (`🔘 ກຳລັງຮຽນ` / `🎓 ຈົບ` / `📋 ທັງໝົດ`) drive the cohort filter via `DB.GetHistoricalClassRoster(..., statusFilter)`. Actions implicitly skip non-active rows.

### Student Score History (`ScoreHistoryPage`, idx 13)

**Page = class-roster picker.** Filter bar: Year + Grade + Room + 🔍 search. Below the filters sits a status-filter bar with three mutex radios + the global class-history button:

| Radio | Filter | Default |
|---|---|---|
| `🔘 ກຳລັງຮຽນ` | `Students.Status='ກຳລັງຮຽນ'` (active only) | ✅ |
| `🎓 ຈົບ` | `Students.Status='ຈົບ'` (graduated only) | |
| `📋 ທັງໝົດ` | no status filter (ກຳລັງຮຽນ + ຈົບ + ອອກ) | |

Changing a radio re-fires `ReloadRoster()` which re-queries `DB.GetHistoricalClassRoster(year, grade, room, statusFilter)` with the matching status. Year/Grade/Room dropdowns are NOT touched. The per-window report-type dropdown (in `StudentHistoryWindow` / `ClassHistoryWindow`) is independent of this page — its `Settings.score_history_report_type` is never reset.

Alongside the radios: **one global** `📚 ປະຫວັດທັງຫ້ອງ` button (disabled until all three filters are picked) — clicking it opens `ClassHistoryWindow` for the current selection. The per-row `📚` button was removed because the class report is identical for every student in the same class. When all three filters have a value, the page loads `DB.GetHistoricalClassRoster(year, grade, room, statusFilter)` and displays the cohort in a DataGrid with columns:

| ລະຫັດ | ຊື່ນັກຮຽນ | ເພດ | ສະຖານະ | 📖 ປະຫວັດສ່ວນຕົວ |

The single per-row button opens `StudentHistoryWindow(sid)`. Search filters client-side on `ລະຫັດ` OR `ຊື່ນັກຮຽນ`. No history data loads until a button is clicked.

**Historical cohort logic (`DB.GetHistoricalClassRoster`):** UNION of two sources:
1. GradeHistory rows whose smallest `AcademicYear > Y` row has `FromGrade=G` and `ClassRoom=R` — that row describes year Y for the student.
2. Students still active in Y (have any `Enrollments.AcademicYear=Y`) whose current `GradeLevel=G` + `ClassRoom=R`, AND no later GradeHistory exists (so the current values still describe Y).

This handles graduated students, promoted students, and current students uniformly.

#### `StudentHistoryWindow` — opened by 📖   (= **Individual Monthly Report** export)

On-screen body still shows the historical multi-year breakdown for browsing:

```
📅 ປີ <year>  ·  ຊັ້ນ <grade>  ·  ຫ້ອງ <room>
   🗓 ເດືອນ 9 / 10 / 11 / 12  (per-subject monthly /10)
   📘 ສະຫຼຸບພາກຮຽນ 1     ສະເລ່ຍ / ຄະແນນເສັງລວມ / ອັນດັບ / ລະດັບ
   🗓 ເດືອນ 2 / 3 / 4 / 5
   📗 ສະຫຼຸບພາກຮຽນ 2     same shape
   📒 ສະຫຼຸບປະຈຳປີ        ສະເລ່ຍພາກ 1 / ສະເລ່ຍພາກ 2 / ສະເລ່ຍປະຈຳປີ / ອັນດັບ / ລະດັບ
```

**Export bar at the bottom** has `ປີ` dropdown + `ປະເພດລາຍງານ` dropdown (single 11-option selector from `HistoryReportCatalog.Items`) + `✓ <label>` status text + `📋 Excel` + `📄 PDF` + Close.

The 11 report-type options (in academic-year order):
| # | Label | Kind | Dispatches to |
|---|---|---|---|
| 1 | ເດືອນທີ 1 — ກັນຍາ | `M9` | `RenderIndividualMonthlyXlsx(sid, year, 9, …)` |
| 2 | ເດືອນທີ 2 — ຕຸລາ | `M10` | `… month=10` |
| 3 | ເດືອນທີ 3 — ພະຈິກ | `M11` | `… month=11` |
| 4 | ເດືອນທີ 4 — ທັນວາ | `M12` | `… month=12` |
| 5 | 📘 ສະຫຼຸບພາກຮຽນ 1 | `S1` | `RenderIndividualSemesterXlsx(sid, year, 1, …)` — Scores.TotalScore + EvaluationScores(SEM1) |
| 6 | ເດືອນທີ 5 — ກຸມພາ | `M2` | `… month=2` |
| 7 | ເດືອນທີ 6 — ມີນາ | `M3` | `… month=3` |
| 8 | ເດືອນທີ 7 — ເມສາ | `M4` | `… month=4` |
| 9 | ເດືອນທີ 8 — ພຶດສະພາ | `M5` | `… month=5` |
| 10 | 📗 ສະຫຼຸບພາກຮຽນ 2 | `S2` | `RenderIndividualSemesterXlsx(sid, year, 2, …)` |
| 11 | 📕 ສະຫຼຸບປະຈຳປີ | `A` | `RenderIndividualAnnualXlsx(sid, year, …)` — mean of Sem1+Sem2 TotalScore per subject; CHA1/LAB1 from EvaluationScores(ANNUAL) |

All variants use `Templates/ລີພອດຄະແນນບຸກຄົນ.xlsx` exactly. Historical grade/room come from `GetStudentHistoryYears` — never from `Students.GradeLevel/ClassRoom`. PDF conversion via `ReportPage.ConvertXlsxToPdfViaExcel` (Excel COM, byte-identical render). The picked option persists across windows + sessions via `Settings.score_history_report_type` (shared with `ClassHistoryWindow`).

#### `ClassHistoryWindow` — opened by 📚   (= **Classroom Monthly Summary Report** export)

On-screen body for browsing (per-month grids + sem/annual summary tables — same as before). **Export bar at the bottom** has the same `ປະເພດລາຍງານ` dropdown (`HistoryReportCatalog.Items`) + status text + `📋 Excel` + `📄 PDF` + Close. Dispatch:

- `M{2-5, 9-12}` → `RenderClassMonthlyXlsx(year, grade, room, month, roster, …)` — uses `ໃບຄະແນນ.xlsx` Sheet 1
- `S1` → `RenderClassSemesterXlsx(year, grade, room, 1, roster, …)` — uses Sheet 2 (sem-1 layout)
- `S2` → `RenderClassSemesterXlsx(year, grade, room, 2, roster, …)` — uses Sheet 3 (sem-2 layout)
- `A` → `RenderClassAnnualXlsx(year, grade, room, roster, …)` — uses Sheet 4. Per-subject academic = mean of Sem1+Sem2 TotalScore; CHA1/LAB1 from EvaluationScores(ANNUAL)

All variants:

- Rows pre-sorted by descending academic sum so the template's `RANK.EQ` formula renders rows in rank order
- Academic-subject columns hold the values; CHA1 / LAB1 columns are written but sit outside the SUM/AVG/RANK formula ranges
- Tie ranks follow `RANK.EQ` (1224-style); students with any subject `< pass_score` render as `"ຕົກ"` per the template's `IF(COUNTIF(<range>,"<5")>0, "ຕົກ", …)` formula
- Cohort = `DB.GetHistoricalClassRoster(year, grade, room)` — graduated + promoted + current students all qualify
- Sem/Annual variants pull from `Scores.TotalScore` + `EvaluationScores` (SEM1/SEM2/ANNUAL contexts)

#### `DB.SafeFileName(string)`

Strips Windows-invalid filename characters (`\ / : * ? " < > |` → `_`), drops control chars, trims trailing dots/spaces, and re-prefixes reserved DOS device names (CON, PRN, AUX, NUL, COM1-9, LPT1-9 — including extension-suffixed forms like `con.txt`). Used by every export's default filename. Also re-exported as `ReportPage.SafeFileName` for convenience inside view code.

**Performance:** Page load = 0 history queries. Roster reload (Year/Grade/Room change) = 1 query. Opening `StudentHistoryWindow` = `years × (1 monthly + 2 sem summaries + 1 annual)` queries (for browse) + 0 export queries. Opening `ClassHistoryWindow` = 1 roster + 8 monthly grids + 2 sem summaries + 1 annual (for browse). Export adds 1 roster + 1 bulk per-(student × subject) query.

### Backup & Restore (`SettingsPage` → Backup tab)

- Backup: `File.Copy(sis_lao.db → user-chosen path)`.
- Restore: confirmation prompt → `File.Copy(chosen → sis_lao.db)` → reload settings → restart the app.
- DB uses default rollback journal (not WAL), so single-file copy is safe.
- Test data at `bin\Debug\net8.0-windows\sis_lao.db` is independent from the installed-app DB at `%LOCALAPPDATA%\Programs\StudentSIS_Lao\sis_lao.db`.

### Academic-year management (`AcademicYearPage`, idx 11, admin)

Redesigned UI: top header card (📅 current year · 📘 semester · 🟢 active badge · 4 global actions) → year-list DataGrid (7 columns + per-row `📊 ສະຖິຕິ` button) → footer system-wide stats card.

**Four global actions:**
- **➕ ເພີ່ມປີໃໝ່** — opens `AcademicYearFormWin`; calls `DB.CreateAcademicYear`. Does NOT auto-promote students or change current year.
- **➡️ ຂຶ້ນປີຕໍ່ໄປ** — computes `DB.NextYearString(DB.CurrentYear)`, auto-creates the row if missing, then `DB.SetCurrentAcademicYear` switches the system. Semester resets to 1.
- **🔄 ຕັ້ງເປັນປະຈຸບັນ** — on selected row, calls `DB.SetCurrentAcademicYear` (single-current invariant + Settings sync).
- **🗑️ ລຶບປີ** — opens the unified `AcademicYearDeleteWindow` (replaces the old MessageBox + `ConfirmTypeWindow` two-step). Shows row-by-row counts of every cascade-deletable table; force-delete button enables only after the user types `DELETE`. Cascade runs in ONE transaction:  Enrollments → EvaluationScores → AttendanceRecords → GradeHistory → Students(year) → AcademicYears(year). Current year always refused.

**Per-row `📊 ສະຖິຕິ` button** opens `AcademicYearStatsWindow` — read-only popup with student counts (total / active / graduated / withdrawn / distinct classrooms) and score-table counts (enrollments / scores / monthly / evaluations / attendance / history) for that year.

**Consistency:** the page NEVER issues custom `UPDATE Settings` / `UPDATE AcademicYears.IsCurrent` SQL — year switching goes only through `DB.SetCurrentAcademicYear`. The old `ConfirmTypeWindow` was deleted (dead code after the delete-flow rewrite).

---

## 11. Login & user management

| | |
|---|---|
| `LoginWindow.xaml` | Empty fields by default (no pre-fill, no hint box). `Loaded` handler defensively clears + focuses username. |
| `LoginWindow.xaml.cs` | No auto-login. Submit only on button click or Enter in password field. |
| `DB.Login(user, pass)` | Plain `SELECT ... WHERE Username=@u AND Password=@p AND IsActive=1`. Plaintext compare (intentional — single-school offline). |
| `UsersPage` (idx 10) | Add / Edit / Toggle-active. Read-only `DataGrid`. |
| `UserFormWin` | Sections: profile (Username, FullName, Role) + password (new + confirm, optional for edit). Validation, success message, post-save round-trip read-back that runs the same `WHERE Password=@p` query the production `Login()` uses. |

**Default seeded users** (DB init): `admin/admin1234`, `teacher1/teacher1234`. Documented in the user manual as needing immediate password change after install.

---

## 12. File / folder layout

```
StudentSIS_Lao\
├─ App.xaml, App.xaml.cs                  (entry point — sets global font fallback chain)
├─ StudentSIS.csproj                      (single project; <Version> here)
├─ StudentSIS.sln
├─ sis_lao.db                             (NOT shipped — created on first run)
├─ CONTEXT_REFERENCE.md / context_reference.md  (this file — same file, Windows is case-insensitive)
│
├─ Data\
│  ├─ DB.cs                               (static facade — schema + all SQL + helpers)
│  └─ ExcelImport.cs                      (Excel template gen + parse/validate/UPSERT
│                                          for both score-entry pages; shared by tests)
│
├─ Views\
│  ├─ MainWindow.xaml + .cs               (sidebar + page host)
│  ├─ LoginWindow.xaml + .cs              (empty fields, no hints)
│  ├─ DashboardPage.xaml + .cs            (idx 0)
│  ├─ StudentsPage.xaml + .cs             (idx 2) + StudentFormWindow
│  ├─ ScoresPage.xaml + .cs               (idx 5)
│  ├─ ReportPage.xaml + .cs               (idx 7 — all 7 report builders + helpers)
│  └─ RemainingPages.cs                   (BIG file — ClassHub, Enrollment, BatchEnroll,
│                                          MonthlyScores, Promotion, Subjects, Users + form,
│                                          AcademicYear (+ form, Delete window, Stats window), Settings,
│                                          and the small H/ classes for UI builders)
│
├─ Helpers\
│  └─ Converters.cs                       (LevelToBg/Fg converters for Level chips)
│
├─ Styles\
│  ├─ Colors.xaml                         (solid colour palette — flat design, no gradients)
│  └─ Controls.xaml                       (button styles, card style, etc.)
│
├─ Templates\                             (4 files — ship next to EXE via PreserveNewest)
│  ├─ ໃບຄະແນນ.xlsx
│  ├─ ໃບສັນຍາ.docx
│  ├─ ລີພອດຄະແນນບຸກຄົນ.xlsx
│  └─ ລາຍງານປະຫວັດນັກຮຽນ.docx
│
├─ installer\                             (build pipeline — see §13)
│  ├─ StudentSIS_Lao.iss                  (Inno Setup script)
│  ├─ build.cmd                           (publish + compile installer)
│  ├─ generate-manual.cmd                 (one-click user manual)
│  ├─ User-Manual.pdf                     (18 pages, Lao, ~4 MB)
│  └─ README.md
│
└─ tests\
   └─ IntegrationTests\                   (console app, net8.0 — links Data\DB.cs directly)
      ├─ IntegrationTests.csproj
      ├─ Program.cs                       (174-assertion test suite + arg dispatcher)
      ├─ UserManual.cs                    (manual generator — OpenXml + Word/LibreOffice converter)
      ├─ Inspect.cs                       (template inspection helpers)
      └─ Repack.cs                        (one-off docx repack utility)
```

Notable file: `Views/RemainingPages.cs` is intentionally one big file holding many pages. Splitting it would help maintenance but isn't necessary.

---

## 13. Build / publish / installer pipeline

### Day-to-day build

```cmd
dotnet build StudentSIS.csproj
```

The 6 NU1701 warnings (iTextSharp / BouncyCastle from .NET Framework) are pre-existing and expected — ignore them.

### Integration test suite

```cmd
dotnet run --project tests\IntegrationTests
```

**490 assertions, all currently passing.** Covers schema integrity, calculation helpers, CHA1/LAB1 contract, EvaluationScores round-trip, ranking semantics, backup/restore fidelity, force-delete cascade, Excel/docx template anchor cells, auto-enrollment idempotency, batch enrollment filters, user-management CRUD + login round-trip, the score-history contract (historical grade/room survives promotion; graduated students' history still queryable; CHA1/LAB1 excluded from semester aggregates), and Excel score import end-to-end (template magic, scope mismatch rejection, validation cascade, CHA1/LAB1 routing, UPSERT, and `Scores.MidScore` recompute after monthly import).

### Build the installer

```cmd
installer\build.cmd
```

Prereqs: .NET 8 SDK + Inno Setup 6.x.

Pipeline:
1. `dotnet publish -c Release -r win-x64 --self-contained true` → `publish\` (~180 MB, 271 files)
2. Locates `ISCC.exe` (PATH + common install dirs)
3. Compiles `installer\StudentSIS_Lao.iss` → `installer-output\StudentSIS_Lao_Setup_v1.0.0.exe`

Installer defaults to per-user install at `%LOCALAPPDATA%\Programs\StudentSIS_Lao` (no UAC). User can elevate to install for all users via Inno's dialog. **Never touches `sis_lao.db`** — preserved across upgrades and uninstalls.

### Regenerate the user manual

```cmd
installer\generate-manual.cmd
```

Generates a 14-chapter Lao .docx via OpenXml, then converts to PDF via Word COM (or LibreOffice headless fallback). Output: `installer\User-Manual.pdf`.

**Why not iTextSharp directly?** iTextSharp 5 doesn't do OpenType shaping for Lao combining marks; the result is unreadable. Word handles complex scripts natively, so we route through .docx.

---

## 14. Reused integration-test invocation flags

The test program also serves as a dev utility — pick the subcommand:

| Command | Effect |
|---|---|
| `dotnet run --project tests/IntegrationTests` | Runs the full 174-test suite |
| `... -- inspect <path.xlsx>` | Dumps cells, merges, page setup, column widths |
| `... -- inspect-docx <path.docx>` | Lists `{{TOKEN}}` placeholders found in a Word file |
| `... -- inspect-pdf <path.pdf>` | Page count, char counts, Lao-char counts |
| `... -- repack <dir> <out.docx>` | Rebuild a docx from an unpacked folder (used when patching templates) |
| `... -- manual <out.pdf-or-docx>` | Generate the user manual |

---

## 15. Conventions to follow when adding code

1. **Match the surrounding style.** RemainingPages uses dense one-liners; ReportPage and newer code uses spacious multi-line — don't mix the two within a method.
2. **All MessageBox calls take title + icon.** `MessageBox.Show(text, title, OK, Information/Warning/Error)`. Don't use the single-arg overload.
3. **Inline status feedback for in-flow saves** (e.g. `_info.Text = $"✅ ..."`); MessageBoxes only for destructive confirmations or errors.
4. **Lao first** for all user-visible strings. Code identifiers and DB column names stay English.
5. **Year mutations go through `SetCurrentAcademicYear`** — never `SaveSetting("current_year", ...)` directly.
6. **CHA1/LAB1 must never be averaged** — every new aggregate query must include `AND sub.SubjectCode NOT IN ('CHA1','LAB1')`. Add an integration test for any new aggregate.
7. **Database writes that include user data should use parameterised SQL** (`("@p", value)` tuples), not string interpolation.
8. **Run the test suite + smoke-launch the app** before declaring a feature done. Build success alone is not enough; check 174/174 still passes.
9. **Don't write comments that explain what code does** — names should do that. Only comment the *why* when it's non-obvious.

---

## 16. Known things that work "by design" and aren't bugs

| Behaviour | Reason |
|---|---|
| Login uses plaintext password compare | Single-school offline app; threat model is "shared Windows account at school", not network attacks |
| Templates have Lao filenames on disk | They work the same on Windows; the `<None Update>` paths in csproj reference them by exact name |
| `GetClassRanking` doesn't push failed students last | Sort is purely by avg desc; the "failed → ‘ຕົກ’" semantics live in the report-layer (`ShowScorePreview` / `GenClassSummary`) |
| Uninstall preserves `sis_lao.db` | Safer default — a casual uninstall must not lose years of student records |
| iTextSharp gets NU1701 warnings | Pre-existing — package is .NET Framework, used via compatibility shim. Build works fine. |
| `Templates/ລາຍງານປະຫວັດນັກຮຽນ.docx` is patched (15.6 KB) vs original (~18 KB) | We fixed the truncated `{{MOTHER_PROVINCE}}` token by repacking |
| No "Remember me" on Login | Removed by explicit user request for security |
| Pass-score field removed from Settings UI | School-wide grading rule; shouldn't change mid-year. Underlying `pass_score` Setting still exists and drives all comparisons. |
| Class Summary + Transcript reports removed | Cleaned up by user request — only 7 report types remain in the dropdown |

---

## 17. Recent significant changes (reverse chronological — keep this trimmed to last ~10)

1. **Score-entry pages: auto-load on filter change, ໂຫຼດ button removed + 3-button dirty-prompt + save writes to loaded tuple**: `ScoresPage` (idx 5) + `MonthlyScoresPage` (idx 6) — the manual `🔄 ໂຫຼດ` button is gone. `_autoReload` starts at `true` so every filter change triggers `LoadRoster` immediately via the existing `OnFilterChanged` path. Constructor now kicks an initial `LoadRoster(manual: false)` if there's no `DB.NavGrade` handoff and a subject is selected, so teachers see the default class roster on page open. F5 still works as a keyboard-shortcut manual reload. The old plain YesNo `ຈະຍົກເລີກບໍ?` MessageBox (which used English-locale "Yes"/"No" labels and confused "cancel edits" with "cancel filter change") was replaced by the new **`SaveConfirmDialog`** — a custom 3-button window (💾 ບັນທຶກ · ບໍ່ບັນທຶກ · ຍົກເລີກ) that returns a `SaveConfirmResult` enum. Save routes to `BtnSave_Click` / `SaveAll` and only proceeds if `_dirty` cleared; DontSave discards edits and proceeds; Cancel keeps edits and blocks the filter change. **Save methods now write to the LOADED tuple** (`_loadedYear` / `_loadedSem` / `_loadedSubCode` / `_loadedMonth` / `_loadedContext`) instead of reading the current combos, so the SaveConfirmDialog Save path can safely save mid-transition without dumping rows into the wrong bucket. MonthlyScoresPage gained a new `_loadedSubjectCode` field for this. The old "ຕົວກອງປ່ຽນແລ້ວ ຫຼັງຈາກໂຫຼດຂໍ້ມູນ … ກົດ 'ໂຫຼດ' ໃໝ່ ກ່ອນບັນທຶກ" filter-drift MessageBox was removed from both `BtnSave_Click` and `SaveAll` / `SaveEval` — it referenced a button that no longer exists and is redundant now that save uses `_loadedX` as the source of truth. Route selection in `SaveAll` also switched to `_loadedContext != ""` (was `IsEvalMode` reading current combo). `InvalidateRoster` hint text updated. Footer hint on ScoresPage updated. All 498 integration tests still pass.
2. **App icon (gear + graduation cap)**: `Assets/app-icon.png` (1024×1024 source) converted to multi-resolution `Assets/app-icon.ico` (16/32/48/64/128/256, ~40 KB) via a PowerShell + `System.Drawing` script. Hooked at three layers: (a) `.csproj` `<ApplicationIcon>` embeds it in `StudentSIS_Lao.exe` (Explorer / taskbar / Alt+Tab); (b) `MainWindow.xaml` + `LoginWindow.xaml` set `Icon="/Assets/app-icon.ico"` for the title-bar; (c) `installer/StudentSIS_Lao.iss` `SetupIconFile=..\Assets\app-icon.ico` for the setup wizard + Add/Remove Programs. Shortcuts in `[Icons]` inherit from the exe. Also added `<None Update="Assets\app-icon.ico">PreserveNewest` + `<Resource Include="…">` so the icon ships next to the exe and resolves at runtime.
3. **PromotionPage — next-year guardrail before Promote / Repeat**: both flows now call `EnsureNextYearExists(selected)` before opening `PromotionConfirmWindow`. Computes `dstYear = NextYearString(MostCommon(srcYears))`; if it's missing from the `AcademicYears` registry, shows a Yes/No prompt (`ຕ້ອງເພີ່ມປີການສຶກສາ ‘ປີໜ້າ’ ກ່ອນ …`) → on Yes, opens `AcademicYearFormWin` **prefilled with the required year** → re-verifies after save (user may have edited the textbox) → cancels the promotion if still missing. `AcademicYearFormWin` gained an optional `prefillYear` constructor param for this handoff. `PromotionConfirmWindow`'s year dropdown was tightened to list ONLY registered years (removed the old `Insert(0, dstYear)` fallback that used to sneak an unregistered year into the picker). All 498 integration tests still pass.
4. **PromotionPage — ຊ້ຳຊັ້ນ (repeat grade) bulk action**: new `🔁 ຊ້ຳຊັ້ນ (ທີ່ເລືອກ)` button on the action toolbar (SecondaryButton style, next to Promote/Graduate) for retaining students who didn't pass. Flow reuses `PromotionConfirmWindow` — `OpenPromoteDialog` gained an `isRepeat` param that flips defaults: `dstGrade = srcGrade` (same grade), `dstYear = NextYearString(srcYear)`, `dstRoom = srcRoom` (editable). `PromotionConfirmWindow` gained an `action` string param that drives the window title (`ຢືນຢັນການຊ້ຳຊັ້ນ` vs `ຢືນຢັນການຂຶ້ນຊັ້ນ`), the header text, and hides ‘ຈົບ’ from the destination-grade dropdown when repeating. Writes `Students(GradeLevel unchanged, ClassRoom=dstRoom, AcademicYear=dstYear, Status='ກຳລັງຮຽນ')` and `GradeHistory(FromGrade=ToGrade=srcGrade, Note='ຊ້ຳຊັ້ນ')`. Log tag `Repeat` (vs existing `Promotion`). All 498 integration tests still pass.
5. **ClassHubPage redesigned — palette-consistent, more info, more entry points**: rebuild aligns the page with the app design system: uses the shared `Card` / `SectionTitle` / `PageTitle` styles and palette brushes (`PrimaryBrush`, `SuccessBrush`, `SecondaryBrush`, `InfoBrush`, `NeutralBrush`, `PrimaryLightBrush`, `TextSecondary/Muted…`) via `SetResourceReference` + `Application.Current.FindResource` — no more hard-coded `Color.FromRgb`. View 1 (grade picker): each ມ.1–ມ.4 card now shows total students + active/graduated pills + rooms-in-use line + hover accent border. View 2 (hub): breadcrumb-style header (`ໜ້າຫ້ອງຮຽນ / ຊັ້ນ ມ.4 · ຫ້ອງ 1 · ປີ ...`), 4-card stats row (Total · Active · Subjects enrolled · Current sem) that live-refreshes when ຫ້ອງ/ປີ change, and 2-column action panels. Left column gained a `ບັນທຶກຄະແນນພາກຮຽນ / ປະຈຳປີ` section with **📘 ພາກ 1** · **📗 ພາກ 2** · **📕 ປະຈຳປີ (CHA1/LAB1)** buttons that route to ScoresPage (idx 5). Right column gained **📚 ປະຫວັດຄະແນນທັງຫ້ອງ** button that routes to ScoreHistoryPage (idx 13) — replaces the old "ສະຫຼຸບຄະແນນປະຈຳພາກ" button that pointed at ReportPage's now-removed score reports. New nav-context handoff: `DB.NavSemester = 3` sentinel puts ScoresPage into ANNUAL mode (`CmbSem.SelectedIndex = 2`). Added `ApplyNavContext` to ScoresPage + ScoreHistoryPage (they were missing pickup — ClassHub used to set `DB.NavGrade` but neither page consumed it). All 498 integration tests still pass.
6. **ScoresPage — ປະຈຳປີ (ANNUAL) entry for CHA1/LAB1**: `CmbSem` gained a third option `ປະຈຳປີ` (SelectedIndex=2). Picking it filters the Subject combo to CHA1/LAB1 only (academic subjects have no manual annual — their annual is derived from Sem1+Sem2 totals) and routes reads/writes through `EvaluationScores(ANNUAL, code)` instead of `SEM{N}`. New `IsAnnualMode` + `Sem_Changed` + `ApplyImportButtonsEnabled` helpers; `ReloadSubjects` gains an `annualFilter` clause; `LoadRoster` + `BtnSave_Click` derive `evalCtx` / `ctx` = `IsAnnualMode ? "ANNUAL" : $"SEM{sem}"`; `ColEval.Header` becomes `ຄະແນນປະຈຳປີ (/10)` when annual; status text says `ປະຈຳປີ` instead of `ພາກ 3`. Import/Download buttons disabled in annual mode (SIS_LAO_SEMESTER_v4 template still covers ພາກ 1/2 only). All 498 integration tests still pass.
7. **History window: per-subject tables for semester + annual views, single dropdown drives both display + export**: `StudentHistoryWindow` + `ClassHistoryWindow` — the report-type dropdown now does DOUBLE DUTY (filter the on-screen body AND set the export target), eliminating the "why doesn't the export match what I see" surprise. Picking S1/S2/A no longer shows a small summary card — it shows a per-subject table (student view) or a subject-per-column grid (class view) matching the monthly-view shape, so teachers see every subject's score at a glance. New DB helpers: `GetClassSemesterGrid` · `GetClassAnnualGrid` · `GetHistoryAnnual`. Also fixed `GetHistoryMonthly` to UNION-in CHA1/LAB1 rows (they were silently dropped by an INNER JOIN on `MonthlyAssessments` — a table CHA1/LAB1 never touch). Rebuild is in-memory (no re-query on filter change). 498 tests still pass with new CHA1/LAB1 monthly assertions added.
8. **ScoresPage redesigned class-wide × one-subject** (mirrors `MonthlyScoresPage`): the per-student-pick layout (one row per subject for one student) is gone. New layout: pick Year · Grade · Room · Status · **Subject** · Sem → the grid loads the whole class (~30 rows) for that one subject → fill the `ສອບເສງ` column down → Ctrl+S saves all rows in one transaction → switch Subject combo to grade the next subject. **30× fewer saves per class per semester** (14 saves instead of 30 × 14 = 420). Keyboard-first: **Enter on the score cell commits + moves DOWN** to the next row, Ctrl+S saves, F5 reloads. CHA1/LAB1 swaps the grid to a single `ຄະແນນ (/10)` column (writes `EvaluationScores(SEM{N})`); academic writes `Scores.FinalScore` while preserving Mid (owned by `RecomputeMidFromMonthly`). Removed: `CmbStudent` + Prev/Next nav + `StudentBanner` + per-student empty state + separate `CmbImportSubject` (import now uses the main Subject combo). New `SemesterRow` / `SubItem` models replace `LaoScoreRow`. Per-student view moved to ScoreHistoryPage (📖 button). All 490 integration tests still pass — data-layer invariants unchanged.
9. **Report templates replaced with school's edited version + signature-row preservation**: `Templates/ໃບຄະແນນ.xlsx` and `Templates/ລີພອດຄະແນນບຸກຄົນ.xlsx` updated from bin/ copies. Class report gained 3 new header rows (national name / motto / school name) pushing data start row 3 → row 6 and putting a signature row (ອຳນວຍການ / ວິຊາການ / ຄູປະຈຳຫ້ອງ) at row 10 — inside the range that `FillSheet` strips. `MapSheet` + `FillSheet` rewritten to detect `HeaderRow` dynamically + snapshot/re-emit signatures below students. Individual report gained typo fix (`ລິດ` → `ລັດ`) and 4 signature lines at row 27.
10. **Semester import template simplified — Midterm column dropped**: 4 visible columns (`ລຳດັບ · ລະຫັດ · ຊື່ · ຄະແນນສອບເສງພາກຮຽນ (/10)`). MidScore stays owned by `RecomputeMidFromMonthly`. Metadata moved to hidden col **Y**, magic `_v4`.

---

## 18. Quick-start cheatsheet for cold pickup

```cmd
:: Where am I?
cd "C:\Users\NITRO V15\OneDrive\Desktop\StudentSIS_Lao"

:: Build
dotnet build StudentSIS.csproj

:: Run the 174-test integration suite
dotnet run --project tests\IntegrationTests

:: Launch the app
bin\Debug\net8.0-windows\StudentSIS_Lao.exe

:: Build the user installer (.exe)
installer\build.cmd

:: Regenerate the user manual (PDF)
installer\generate-manual.cmd
```

**Default credentials (fresh install):** `admin` / `admin1234` (change immediately).

**Most-touched files when adding a feature:**
- New page in sidebar → `MainWindow.xaml.cs` (`_pages` array) + new class in `Views\RemainingPages.cs`
- New report → `ReportPage.xaml` (combobox item) + `ReportPage.xaml.cs` (`BtnGen_Click`/`BtnPdf_Click`/`RefreshPreview` cases + new `GenX` method)
- New DB column → migration block in `DB.cs Initialize()` + integration test
- New calculation rule → helper in `DB.cs` + integration test covering boundary cases

---

## 19. Maintaining this file

This document is a working tool, not history. When you make changes that affect any of the sections above:

1. Update the affected section **in place**.
2. Bump the "Last updated" date at the top.
3. Add a one-line entry to §17 "Recent significant changes" — and trim the bottom of that list to keep it at ~10 entries.
4. If you delete a feature, also delete its row from the relevant table — don't leave stale rows with "removed" notes.

If something in this file disagrees with the code, **the code wins**. Fix this file to match.
