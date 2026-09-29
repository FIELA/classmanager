# 班级助理 (Class Assistant)

A local web tool for subject teachers: manage seating charts for several classes, draw students at random during class, record points on the spot, and track each student's seat history and classroom performance. Open the page by double-clicking it — no installation required. The user interface is in Chinese.

## Features

- **Multiple classes**: one class is shown at a time, switch with one click. Rosters can be imported from Excel or CSV, or pasted directly; name, gender, score, class and similar columns are recognized automatically. Within one class, rows with the same name are merged into one student; if a name also appears in another class, the teacher is asked to confirm before importing.
- **Custom seating layouts**: set the number of rows and the maximum columns; apply seat groupings such as `2+2+2` or `3+2+3` to a range of rows, or adjust aisles and unavailable seats row by row.
- **Several seating methods**: random (optionally pairing boys and girls as deskmates), choose seats in score order, automatic placement by score, front/back and left/right rotation, and keep the previous chart. Individual students can be pinned to a seat before generating.
- **Student view / teacher view**: both are generated from the same data; the teacher view is rotated 180°. Boys' names are blue, girls' red. Click to fine-tune seats.
- **Fast gender marking**: choose "mark girls" (or "mark boys"), click students one by one, mark the rest with one click, and undo. Works from the roster when there is no seating chart yet.
- **Random draw**: draw 1–6 students or a custom number; drawn students are highlighted on the chart and shown prominently above it. Optionally no repeats within a lesson.
- **Group recitation**: each class has one course representative and a fixed sample of 5–10 students. Every week 5 sample students recite to the representative and each then leads a group. The representative and 4 students drawn from all non-representative students recite to the teacher; a student may also be a sample group leader. The remaining students are split evenly into five groups. An animation reveals, in order, the list reciting to the teacher, the list reciting to the representative with the group leaders, and each group's members. After confirmation the result is finalized per week, and past versions stay viewable.
- **Classroom points**: click students on the chart to add or deduct points, with multi-select and undo. Point items are customizable.
- **Performance and seat history**: view class performance and per-student details by week; view each student's past seats and front/middle/back and left/center/right distribution.
- **Excel import and export**: rosters, seating charts (student + teacher view) and classroom performance can all be exported. Edits made in Excel are read back into the page automatically.

## Advantages

- **Local and private**: nothing is uploaded; student data stays only in the folder you choose. The only network access is a clock check against public time services (a plain request that carries no data). Without internet the page works fully offline and uses the computer's own clock.
- **Internet time as the reference**: when online, all dates and record times follow internet time. The page warns when a computer's system clock is off, or when records come from a computer whose clock ran ahead.
- **Plug-and-play from a USB drive**: the whole project can live on a USB drive and be opened directly on classroom computers; works on Win7 (Chrome/Edge 86 or later).
- **Conflict-free sync across devices**: data is stored as "create-only, never modify" record files. With SyncTime, OneDrive or a USB drive shared between computers, no file is ever changed on both sides, and changes from every computer are merged automatically.
- **No lost data**: every change is saved immediately. If the USB drive is pulled out, changes are kept in the browser and saved again when it is plugged back in.
- **Two-way Excel**: exported workbooks can be edited in Excel/WPS and the edits are read back. Editing an old workbook never overwrites newer data, and editing the same exported workbook several times counts as one continuous edit, not a conflict.

## Quick start

1. Open `班级助理.html` in Chrome or Edge.
2. Click "连接文件夹" (Connect folder), choose the project folder, and allow read/write access.
3. In "班级设置" (Class settings), import a roster and set the seating layout, then return to "座次" (Seating) to generate a chart and confirm it ("确认定版").
4. Before using "分组过关" (Group recitation), open "固定人员设置" on that page, choose the course representative and 5–10 sample students, and save. At least 15 active students are required; group sizes are balanced automatically based on class size and overlapping roles. After drawing, check the candidate result, then click "确认定版". Switch the week or expand "查看历史版本" to see earlier arrangements.

## Development

- Code lives in `辅助资源/assets/`. Run the tests with `node --test 辅助资源/tests/*.test.js`.
- After cloning, enable the privacy hooks once: `git config core.hooksPath .githooks`.
- Student data is stored only in the project's `数据/` and `导出/` folders and never enters the repository. Maintenance conventions are in `AGENTS.md`.
