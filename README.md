# Timetable System

We built a three-stage timetable system. Stage 1 builds a valid weekly timetable with balanced science blocks, section composition that increases distinct classmates, and teacher assignment under department caps. Stage 2 assigns one fixed room per Section, avoids room-time conflicts, respects single/double weeks, and reduces student walking using real student transitions and simulated annealing. Stage 3 renders individual student timetables. The time skeleton is fixed to satisfy the school's hard rules; the optimisation happens in section composition and room assignment.

## Pipeline

Stage 1 — classes. The `Week pattern` sheet fixes which subject families occupy which time slots across the week. `Sections` and `Classes` determine how students are grouped into Sections per subject and level, with the aim of maximising distinct classmates. `Teachers` assigns one teacher per Section under department caps. This stage is already resolved in the input workbook; this repository consumes its output.

Stage 2 — rooms (`assign_rooms.py`).** Reads `Enrollments` from the master workbook. Each Section receives one fixed room for the whole week. Sections may share a room at different times, but no room hosts two Sections in the same slot. Single-week (`二外`) and double-week (`CAS`) rotas are treated as separate occupancy. The objective is the total student walking cost over consecutive lessons, minimised by simulated annealing with a conflict penalty. Distances follow the 21-room layout described under Model.

Stage 3 — student timetable (`student_timetable.py`). Reads the room-assigned master workbook and renders one student's weekly timetable: each cell shows subject, level, room and teacher, with single/double week lessons labelled.

## Files

- `assign_rooms.py` — Stage 2. Room assignment and walking-cost optimisation.
- `student_timetable.py` — Stage 3. Individual student timetable rendering.
- `master_schedule.xlsx` — Input master workbook. Must contain `Enrollments` with `Student / Subject / Level / Section / Teacher / Weekly minutes / When`.
- `requirements.txt` — `openpyxl>=3.1`.

## Setup
pip install -r requirements.txt
Place assign_rooms.py, student_timetable.py and master_schedule.xlsx in the same directory. Keep master_schedule.xlsx closed in Excel while running Stage 2.

Run
Stage 2 — room assignment:

bash
python assign_rooms.py
Writes to output/:

rooms_layout.xlsx — 21×21 distance matrix and Section→Room map

rooms_by_slot.xlsx — room grid per time slot (red = conflict; should be empty)

metrics.xlsx — total walking cost a, random baseline, improvement %, per-student a_i

master_schedule.xlsx — master workbook with the Room column filled in

Completion is confirmed by conflicts=0 and [fill_master_room] 写入 N 行 Room.

Stage 3 — student timetable:

bash
python student_timetable.py S014 output/master_schedule.xlsx
Writes output/students/S014_timetable.xlsx. The second argument must point at the Stage 2 output. If it is omitted, the script falls back to the original workbook, which has no Room column and renders (no room).

Batch:

powershell
1..250 | ForEach-Object { python student_timetable.py ("S{0:D3}" -f $_) output/master_schedule.xlsx }
Model
21 rooms. S1..S18 form a weighted ring with edge weights [1,1,3,1×14,3]; S19–S20=1, S20–S21=1, S19–S21=2; any S1..S18 to any S19..S21 = 10.

Each Section receives one fixed room for the whole week; Sections may share a room at different times, but no room hosts two Sections in the same slot.

Walking cost for consecutive lessons: a = 0 if d = 0, else c + x·ln(d). Total a is summed over all students and consecutive lesson pairs.

c, x are selected by grid search on an 80/20 student split. Stage 2 reports the improvement of a relative to a randomised feasible assignment.

Notes
The time skeleton in Week pattern is fixed. Stage 2 does not move lessons; it only assigns rooms.

Single/double week rotas (二外 单周, CAS 双周) are preserved: both appear in the same slot but are rendered separately in Stage 3.

If the peak number of concurrent Sections exceeds 21, Stage 2 prints a warning. The current data peaks at 13.