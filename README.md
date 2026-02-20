# RetreatPlacer

Automatic room assignment for retreat centers using constraint optimization. Given a list of rooms and a list of attendees with preferences, RetreatPlacer finds the best placement that respects accessibility needs, keeps groups together, and organizes organizations into the same buildings.

Built on [Google OR-Tools CP-SAT](https://developers.google.com/optimization/cp/cp_solver), a constraint-programming solver.

## Quick Start

### Desktop UI (recommended)

Place `RetreatPlacerUI.py` alongside `RetreatPlacer.py`, then:

```bash
pip install customtkinter ortools openpyxl pandas
python RetreatPlacerUI.py
```

### Command Line

```bash
pip install ortools openpyxl pandas
python RetreatPlacer.py RoomMap.xlsx PeopleToPlace.xlsx FilledRoomMap.xlsx
```

Or generate sample data to try it out:

```bash
python RetreatPlacer.py --generate-sample
python RetreatPlacer.py
```

---

## Desktop UI

`RetreatPlacerUI.py` is a cross-platform desktop application built with [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter). It provides a visual interface for the full workflow: loading files, running the solver, reviewing results, making manual adjustments, and saving the output.

### UI Tabs

#### Files Tab

Select your input and output files using the native file browser. As soon as files are selected, a summary panel shows building count, bed capacity, org/group breakdown, and constraint counts (floor 1 needs, bottom bunk needs, attachments).

#### Run Solver Tab

Click **Run Solver** to start the OR-Tools CP-SAT solver. The UI automatically switches to the Console Log tab so you can watch progress. The solver runs in a background thread — the UI stays responsive throughout. A configuration panel shows the current solver weights and settings for reference.

#### Console Log Tab

Live output from the solver appears here, including attach-name resolution, solver progress, and the final debug report. While the solver is running, an animated indicator at the top shows a spinner and elapsed timer. When the solver finishes, the indicator turns green and the UI automatically navigates to the Results tab after a short delay.

#### Results Tab

Displays the placement results in a visual layout organized by building and room.

**Summary cards** at the top show total placed, unplaced, buildings used, and rooms used.

**Building cards** show each building with its rooms laid out in a grid. Each room card displays:

- Room name and occupancy (e.g., `7/10`)
- Filled bunk slots showing the person's name and org
  - **Colored names** (matching the building accent color) with a `▾` indicator = **bottom bunk**
  - **Gray names** with a `▴` indicator = **top bunk**
- Empty bunk slots shown as `── empty ──`

**Unplaced section** appears below the buildings if anyone couldn't be placed, with diagnostic reasons for each person.

**Manual adjustments** — You can move people between rooms after the solver runs:

1. Click any person's name to select them (highlights in yellow)
2. All empty slots change to `┄┄ drop here ┄┄`
3. Click an empty slot to move the person there
4. Click the **Deselect** button or another person to change your selection
5. Unplaced people can also be clicked and placed into empty slots

**Save button** — At the bottom of the Results tab, the **Save Changes to Output File** button writes your manual adjustments back to the xlsx. It is grayed out until you make changes, then turns green. An "Unsaved changes" indicator also appears in the header.

Additional buttons let you open the output file directly or open the containing folder.

---

## Distributing to End Users

There are two ways to distribute RetreatPlacer to people who don't have Python installed.

### Option A: Standalone App (PyInstaller)

Bundle everything into a single executable. Users double-click and go — no Python install needed.

**You must build on each platform** (Windows exe on Windows, macOS app on macOS).

#### Setup (both platforms)

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS:   source .venv/bin/activate
pip install -r requirements.txt
pip install pyinstaller
```

#### Build on Windows

```
build_windows.bat
```

Output: `dist/RetreatPlacer.exe`

Users may see a SmartScreen warning on first run — click "More info" → "Run anyway". To avoid this, code-sign the exe with a certificate (~$200-400/yr).

#### Build on macOS

```bash
chmod +x build_macos.sh
./build_macos.sh
```

Output: `dist/RetreatPlacer.app` (also creates a distributable zip)

Users may need to right-click → Open on first launch (Gatekeeper). To remove this, notarize with an Apple Developer account ($99/yr). If users see "App is damaged", they can run `xattr -cr /path/to/RetreatPlacer.app` in Terminal.

**Apple Silicon vs Intel:** Build on the same chip architecture your users have.

### Option B: Install Scripts (lightweight, requires internet)

Give users the project folder. They double-click one file:

- **Windows:** `install_windows.bat` — checks for Python (opens the download page if missing), creates a hidden virtual environment, installs packages, and launches the app. First run takes 2-3 minutes; subsequent runs start in seconds.

- **macOS:** `install_macos.command` — same flow, double-clickable from Finder.

This approach has a much smaller download size but requires an internet connection on first run.

### Project Files for Distribution

```
RetreatPlacer/
├── RetreatPlacer.py          # solver engine
├── RetreatPlacerUI.py         # desktop UI
├── requirements.txt           # pip dependencies
├── build_windows.bat          # PyInstaller build script (Windows)
├── build_macos.sh             # PyInstaller build script (macOS)
├── install_windows.bat        # one-click install+run (Windows)
└── install_macos.command      # one-click install+run (macOS)
```

---

## Input Files

### RoomMap.xlsx

Defines every room at the retreat center.

| Column | Type | Description |
|---|---|---|
| BuildingName | text | Building name (e.g., "Oak Lodge"). Rooms in the same building are grouped for org cohesion. |
| RoomName | text | Room identifier within a building (e.g., "Room 101"). |
| RoomFloor | int | Floor number: `1` or `2`. |
| #BottomBunk | int | Number of bottom bunks (accessible). |
| #TopBunk | int | Number of top bunks. Use `0` for rooms with only single beds. |

Total room capacity = `#BottomBunk + #TopBunk`.

### PeopleToPlace.xlsx

Lists every attendee and their preferences.

| Column | Type | Description |
|---|---|---|
| FirstName | text | First name. |
| LastName | text | Last name. |
| OrgName | text | Organization name. People in the same org are placed in the same building(s). Leave blank if unaffiliated. |
| GroupName | text | Small group name (e.g., "Team1", "MomLife"). People in the same group are placed in the same room. Leave blank for no group. |
| AttachName | text | Full name ("First Last") of a person to room with. See [Attachment Rules](#attachment-rules) below. |
| RoomLocationPref | text | `1` = must be on floor 1 (accessibility). Anything else = any floor. |
| BunkPref | text | `Bottom` = must have a bottom bunk (accessibility). Anything else = any bunk. |

## Output: FilledRoomMap.xlsx

The output workbook has four sheets:

### Sheet 1: FilledRoomMap

The main placement results, sorted by building → room → bunk.

| Column | Description |
|---|---|
| BuildingName | Assigned building |
| RoomName | Assigned room |
| FirstName, LastName | Person |
| OrgName, GroupName | Organization and group |
| RoomFloor | Floor of assigned room |
| Bunk | `Bottom` or `Top` |
| AttachName | Original attachment value from input |
| AttachResolved | Who it was matched to (after fuzzy resolution) |

### Sheet 2: Unplaced

Anyone who couldn't be placed, with diagnostic reasons explaining why (e.g., "Needs bottom bunk on floor 1 — only 12 such slots exist, likely full").

### Sheet 3: AttachWarnings

Details on every non-exact AttachName resolution: fuzzy matches, rejected matches, skipped group references, and unresolved names. Use this to clean up your source data.

### Sheet 4: Summary

Placement counts, org-by-building distribution, and overall statistics.

---

## How It Works

### Constraint Model

RetreatPlacer formulates room assignment as a constraint satisfaction optimization problem. Each person is assigned to a room (or left unassigned), and the solver maximizes a weighted objective while respecting hard constraints.

**Hard constraints** (must be satisfied):

| Rule | Description |
|---|---|
| Room capacity | Each room's total occupancy ≤ `#BottomBunk + #TopBunk` |
| Bottom bunk capacity | Bottom-bunk-needing people per room ≤ `#BottomBunk` |
| Floor preference | `RoomLocationPref = 1` → person can only be in floor 1 rooms |
| Mutual attachments | If A references B **and** B references A, they must share a room |

**Soft constraints** (optimized via weighted objective):

| Rule | Weight | Description |
|---|---|---|
| Placement | 10,000 | Every person placed earns a large bonus (everyone gets placed first) |
| Group cohesion | 1,000 | Consecutive group members in the same room earn a bonus; different rooms incur a penalty |
| Attachment preference | 800 | One-directional attachments prefer the same room |
| Org building affinity | 200 | Per-person bonus for being in the org's pre-computed preferred building |
| Org cohesion | 100 | Consecutive org members in the same building earn a bonus |

The weight hierarchy ensures placement always comes first, then group cohesion, then attachments, then org cohesion.

### Attachment Rules

AttachName is the most nuanced feature. Here's how it works:

**Mutual pairs (hard constraint):** If Alice lists Bob and Bob lists Alice, they *must* share a room. This is for pairs who absolutely need to be together.

**One-directional references (soft constraint):** If Alice lists Bob but Bob lists someone else (or nobody), it's a strong preference. The solver will try to place them together but won't leave Alice unplaced if it can't.

**Unresolved references:** If the named person isn't in the attendee list at all, no constraint is created — the person is placed normally.

**Group references:** If the AttachName matches a known GroupName or OrgName (e.g., "MomLife"), the person is auto-assigned to that group instead of creating a person-to-person attachment.

### Fuzzy Name Matching

Real-world spreadsheets have typos, nicknames, and inconsistent formatting. RetreatPlacer resolves AttachName values through a 5-step cascade:

1. **Exact match** (case-insensitive): "Jessica Ritchie" → Jessica Ritchie
2. **Nickname expansion**: "Jess Ritchie" → Jessica Ritchie (supports ~20 common nicknames)
3. **Last-name match** with first-name plausibility check and org/group affinity tiebreaking
4. **First-name-only match**: "Heather" → picks the Heather in the same org
5. **Fuzzy matching** (SequenceMatcher) with org/group affinity boost: same-org candidates get a 0.15 bonus per affinity point

Candidates in the same org or group are preferred at every step, preventing false cross-org matches.

### Case Normalization

GroupName and OrgName are normalized to a canonical form (first occurrence wins), so `MomLife`, `Momlife`, and `momlife` are all treated as the same group.

### Org-Building Affinity

Before solving, RetreatPlacer pre-computes which buildings each org should ideally use via greedy bin-packing (largest orgs first → fewest buildings). This gives the solver a strong directional signal that scales with org size.

---

## Solver Details

- **Engine:** Google OR-Tools CP-SAT
- **Time limit:** 600 seconds (configurable in code)
- **Workers:** 8 parallel search threads
- **Model type:** Room-level assignment with per-room capacity constraints

The solver reports whether the solution is OPTIMAL (proven best) or FEASIBLE (good solution found within time limit).

## Requirements

- Python 3.10+
- [ortools](https://pypi.org/project/ortools/) — Google OR-Tools constraint solver
- [openpyxl](https://pypi.org/project/openpyxl/) — Excel file creation
- [pandas](https://pypi.org/project/pandas/) — Excel file reading
- [customtkinter](https://github.com/TomSchimansky/CustomTkinter) — Desktop UI (only needed for `RetreatPlacerUI.py`)

```bash
pip install ortools openpyxl pandas customtkinter
```

## Command Line Usage

```bash
# Basic usage
python RetreatPlacer.py RoomMap.xlsx PeopleToPlace.xlsx FilledRoomMap.xlsx

# Default filenames (if you name your files exactly as above)
python RetreatPlacer.py

# Generate sample data for testing
python RetreatPlacer.py --generate-sample
```

## Template Files

The `templates/` directory contains starter spreadsheets with:

- Column headers and descriptions
- Data validation (dropdown lists for BunkPref, RoomLocationPref)
- Example data (in blue) that you can replace
- An Instructions sheet explaining each column

**Important:** Delete row 2 (the description row) before running the script, or copy only the header row and your data.

---

## Troubleshooting

**Everyone is placed but org cohesion is poor:**
The solver prioritizes placing everyone over org cohesion. If the solution is FEASIBLE (not OPTIMAL), try increasing the time limit in the code (`solver.parameters.max_time_in_seconds`).

**Mutual pairs are unplaced:**
Both people in a mutual pair must fit in the same room. If one needs a bottom bunk on floor 1 and those rooms are full, neither can be placed. Check the Unplaced sheet for diagnostics.

**AttachName matched the wrong person:**
Check the AttachWarnings sheet. If a fuzzy match is incorrect, fix the spelling in PeopleToPlace.xlsx — exact matches are always preferred.

**Solver takes too long:**
The room-level model handles 200+ people well within 5 minutes. For 500+ people, consider increasing the time limit or reducing the number of soft constraints.

**UI won't launch:**
Make sure `RetreatPlacerUI.py` is in the same folder as `RetreatPlacer.py`. The UI imports the solver as a Python module. If CustomTkinter isn't installed, run `pip install customtkinter`.

**"Solver Not Found" error in the UI:**
The UI looks for `RetreatPlacer.py` in the same directory. Verify both files are side by side.

**Windows SmartScreen blocks the standalone exe:**
Click "More info" → "Run anyway". This is a known issue with PyInstaller-built executables that aren't code-signed.

**macOS "App is damaged" or "cannot be verified":**
Right-click the app → Open (first time only). If that doesn't work, run `xattr -cr /path/to/RetreatPlacer.app` in Terminal.

## License

MIT