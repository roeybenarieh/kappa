# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
uv run app.py          # launch the GUI
uv run create_excel.py # regenerate the blank personnel.xlsx template
uv add <package>       # add a dependency
```

## Architecture

Two entry points:

- **`create_excel.py`** — standalone script that generates the blank `personnel.xlsx` template. `app.py` imports `create_workbook()` from here when the user clicks "Get Template".
- **`app.py`** — tkinter GUI with two tabs (Main, Help). All generation and assignment logic lives here.

### Data flow

1. User fills in `personnel.xlsx` (six tables in one sheet):
   - **Soldiers** — cols A–F (Name, Last Name, Personal Number, Phone, Course, Gender)
   - **Commanders** — cols H–J (Name, Last Name, Half Kappa)
   - **Hitnasuyot** — cols L–N (Hitnasut Name, Full Soldier Name, Full Commander Name)
   - **Courses** — cols P–Q (Course, Half Kappa)
   - **Rooms** — cols S–U (Room Number, Gender, Capacity)
   - **Computer Users** — cols W–X (Username, Password) — rows align with Soldiers rows by position; soldier at index i gets `computer_users[i]`
2. `read_excel()` parses each table and validates every row through Pydantic models (`Soldier`, `Commander`, `Hitnasut`, `Course`, `Room`, `ComputerUser`). Invalid rows surface as a single `ValueError` before anything is written. Returns a 6-tuple.
3. `_validate_cross_references()` runs additional checks across tables and returns a `dict[str, list[str]]` of error sections (empty = OK):
   - Duplicate soldier full names (cols A–B)
   - Duplicate personal numbers (col C)
   - Duplicate phone numbers (col D, non-empty only)
   - Soldier course not found in Courses table (col F)
   - Soldier's course half-kappa has no matching commander (col F)
   - Hitnasuyot full soldier name not in Soldiers table (col M)
   - Hitnasuyot full commander name not in Commanders table (col N)
4. `_check_room_capacity()` checks for duplicate room numbers and whether total room capacity per gender covers all soldiers of that gender.
5. `refresh_status()` is called on a 1-second poll (`_poll_excel`) that watches `personnel.xlsx` for file changes. It updates the Generate button colour (green = ready, red = errors) and renders live error labels below the button.
6. Clicking **Generate** runs both validation functions, then:
   - `_assign_rooms(soldiers, rooms)` — CP-SAT room assignment
   - `_assign_commanders(soldiers, commanders, courses)` — CP-SAT commander assignment
   - Writes up to five files into a uniquely-named output directory (`result`, `result(1)`, …):
     - `data.xlsx` — all soldier data combined (name, course, gender, half-kappa, commander, room, hitnasut, API fields, username/password from Computer Users)
     - `enriched_soldiers.xlsx` — soldiers with API-fetched extra fields
     - `hitnasuyot.pptx` — one set of slides per hitnasut group
     - `room_signs.docx` — one page per room with soldier names
     - `name_tags.pptx` — one slide per batch of soldiers

### CP-SAT assignment logic

Both `_assign_commanders` and `_assign_rooms` use OR-Tools CP-SAT with the same two-priority objective:

**Priority 1 (dominant) — no isolated soldiers.**
For each course group with ≥ 2 soldiers, no single soldier may end up as the only member of their course in a given room or commander team. Detected with two bool helpers per (group, slot) pair:
- `at1` = 1 iff at least 1 member of the group is in this slot
- `at2` = 1 iff at least 2 members are in this slot
- `alone = at1 − at2` = 1 iff exactly 1 member is in this slot (isolated)

Penalty weight = `S + 1` (total soldiers + 1), which strictly dominates the maximum possible cohesion gain.

**Priority 2 — maximise fully intact course groups.**
`intact[g][slot]` = bool var that implies all soldiers in group `g` are assigned to `slot`. Objective contribution: `group_size[g]` per intact group.

Combined objective: `maximize  sum(intact * sizes) − (S+1) * sum(alone)`

**Commander-specific constraints:**
- Each soldier is restricted to commanders whose `half_kappa` matches the soldier's course's half-kappa (from the Courses table). Falls back to all commanders if the soldier has no course or no matching half-kappa commander exists.
- **Equal load (hard constraint):** within each half-kappa group every commander must receive between `⌊n/c⌋` and `⌈n/c⌉` soldiers, where `n` = soldiers in that half-kappa and `c` = commanders in that half-kappa. The difference between the most- and least-loaded commander in the same half-kappa is therefore at most 1. This is a hard `model.add(sum >= lo) / model.add(sum <= hi)` constraint that takes precedence over the soft objectives.

**Room-specific constraints:**
- Soldiers are restricted to rooms whose gender matches theirs.
- No room exceeds its stated capacity.
- Results are accumulated per room-index (not room-number) to handle the edge case of multiple rows with the same room number correctly.

**Fallback:** if the solver does not return OPTIMAL or FEASIBLE within 5 seconds, both functions fall back to a deterministic round-robin within half-kappa (commanders) or sequential gender-based fill (rooms).

### Key conventions

- `EXCEL_PATH = "personnel.xlsx"` is the single source of truth for the input file name.
- All generator functions (`_generate_enriched`, `_generate_pptx`, `_generate_room_signs`, `_generate_name_tags`, `_generate_data_xlsx`) are pure: they take data + `out_dir: Path` and return the output path string. Side-effect-free until `btn_generate_all` creates the directory.
- The placeholder API in `_api_fetch` falls back to random mock data on any exception — replace the `API_URL` constant there when the real endpoint is known.
- CP-SAT variable lists must use explicit `list[...]` comprehensions, never generators, to avoid closure bugs with loop variables in `model.add_exactly_one` and sum constraints.
