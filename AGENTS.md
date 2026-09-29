# 班级助理 · Project conventions

- Respond in Chinese by default. Read README.md first. The project was created on 2026-09-26, modelled on the old project "02事务/座次", but it does not migrate old data, does not handle duty rosters, and does not show two classes on one screen.
- Student names are used only inside this project: external records, commit messages and logs must never contain names, rosters or individual seat details.

## Privacy and the repository (public: github.com/FIELA/classmanager)

- Real data lives only in `数据/` and `导出/`; both are ignored by `.gitignore`. `.gitignore` is a whitelist: at the root only the program, docs, `.githooks/` and `.github/` are allowed.
- Enable the hooks once per clone: `git config core.hooksPath .githooks`. `pre-commit`, `pre-push` and `commit-msg` call `辅助资源/tools/check-no-real-data.js`, which blocks the data folders, spreadsheet/data file types, and any student name found in the local `数据/` — in files and in commit messages. Never bypass it (no `--no-verify`).
- CI (`.github/workflows/check.yml`) runs the path/file-type check and the tests on every push. It has no local `数据/`, so the name check only happens in the local hooks.
- Tests and examples use fictitious names only (e.g. "赵测一"); test data is generated in memory and never written to disk.
- The git database lives in `~/.gitdirs/classmanager.git` (the project's `.git` is a text file pointing to it), so that OneDrive and SyncTime never sync git internals.

## Sync safety (core — must not be broken)

Usage scenario: the project sits on a USB drive and classroom computers (Win11 / Win7) open it directly from the drive. On the Mac, SyncTime syncs `备课` ↔ USB drive in both directions, and the Mac folder is also in OneDrive. SyncTime rules: a file changed on both sides is a conflict, deletions propagate to the other side, and changes are detected by modification time and size.

- The page and the tools **only create new files**; they never overwrite, rename, move or delete an existing file. The single exception: the current session's own record file is rewritten during the session (only that session writes it).
- Data = all `*.jsonl` record files under `数据/` (one operation per line) + snapshots under `数据/快照/` (snapshots only speed up loading, and are new files too). `数据/project.json` is written once, when the project is created.
- Every operation carries a unique id and a hybrid-logical-clock timestamp (`store.js`). Merging sorts by timestamp and de-duplicates by id; deletions use tombstones, nothing is really deleted. A malformed operation is skipped on its own (`store.safeApply`) and never blocks loading.
- Two independently saved seating finals or group finals for the same week, two layouts, or two group-member settings are detected as a conflict and the teacher chooses in the UI. Successive read-backs of the same exported workbook are sequential edits, not a conflict. All other fields are last-writer-wins.
- Ignore system/temporary files when reading (`._*`, `~$*`, `*.crswap`, `.DS_Store`, …). File names use only characters allowed by Windows and FAT32, with no trailing space or dot.
- Excel is only an "export": every export is a new file. The hidden sheet `_classmanager` stores the content at export time; read-back compares only the teacher's changes, and the generated operation ids are derived from the file content (if two computers read the same change, the id is the same and it counts once). The page remembers the last read version of each exported file (size + modification time); a failed read is retried after the file changes or the page is reopened.
- Browser storage (localStorage / IndexedDB, keys prefixed `classmanager.v1`) only holds operations not yet written to the folder, one key per session (`classmanager.v1.outbox.<session id>`), cleared after a successful write. The key names differ from the old project.

## Time

- Internet time is the reference. The page queries public time services in order (`store.TIME_SOURCES`, a plain GET with no data; first success wins) at start-up and every 30 minutes; "now" is then local time + offset (clock, dates, record times, file-name stamps). Offline, it falls back to the local clock.
- Warnings: the system clock differs from internet time by more than `store.CLOCK_TOLERANCE` (5 minutes); records are timestamped later than internet time (a computer whose clock ran ahead); offline, local time is more than 1 hour earlier than existing records.
- Snapshots are chosen by the time in their file name: a snapshot later than the reference time is "ahead" and used only when no normal snapshot exists, with a warning. A new snapshot is named with the reference time and always later than the snapshot currently in use, so it is picked next time and snapshots are never rewritten on every load.

## Rosters and same-name students

- Within one class, the same name is the same student: duplicate rows in an import are merged (later rows fill in fields), merging into an existing class matches by name (preferring active students with a lower sequence number), and "add student" refuses a duplicate (or offers to restore a student who left).
- If an imported or added name already belongs to an active student in another class, the teacher must confirm first. `tools/import-roster.js` stops and asks for `--yes` in that case (it prints only counts, never names).

## Implementation

- Entry point `班级助理.html` (the old name `座次管理.html` is still recognized, see `store.ENTRY_FILES`); `辅助资源/assets/`:
  - `core.js` pure rules: dates, layouts and zones (including per-row seat groupings), four seating methods (random / by score / rotation / keep), drawing, group recitation, performance summaries, same-name matching.
  - `store.js` data layer: clock, internet time, merging, record-file I/O, snapshots.
  - `excel.js` roster import, three kinds of export, reading edits back.
  - `app.js` UI; `app.css` styles; `exceljs.min.js` is a third-party library (MIT).
- Coordinates: row r, column c, from the students' point of view; row 1 is nearest the lectern; the teacher view is only rotated 180° for display.
- Compatibility: only syntax and CSS supported by Chrome 86 (Win7 tops out at Chrome/Edge 109); no ES modules (they cannot load under file://). `tests/compat.test.js` checks this.
- Tools: `tools/import-roster.js` (import classes from score sheets), `tools/check-no-real-data.js` (leak check), `tools/node-fs-handle.js` (lets Node read/write disk through the same data layer).

## Tests and acceptance

- `node --test 辅助资源/tests/*.test.js`. When changing rules, write a failing test first. `store.test.js` contains tests that simulate Mac ↔ USB sync under SyncTime's rules; they must keep passing whenever the data layer changes.
- Browser acceptance uses a temporary copy + a local server + the browser's private directory (OPFS): in the console, first create a `数据` subfolder, then call `ClassManagerApp.connectHandle(handle)`. Afterwards clear the private directory, localStorage, IndexedDB and the temporary copy. Never run write tests against a real data folder.
