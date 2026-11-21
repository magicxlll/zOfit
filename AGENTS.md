# zOfit QS Pro — Agent Instructions & Workflow Rules (AGENTS.md)

This file defines the mandatory operating protocol for all AI coding agents working on **zOfit QS Pro**.

## 1. Session Start (Cross-Machine Continuity Protocol)
Whenever the user instructs to read `memory.md` and continue development:
1. Check repository status:
   - If not cloned yet: `git clone https://github.com/magicxlll/zOfit.git .`
   - If cloned: `git fetch origin && git status && git pull origin main`
2. Read project context in order:
   - `memory.md` (Current snapshot, ADRs, Traps, Session log)
   - `tasks/todo.md` (Active tasks)
   - `docs/spec.md` (Business and technical specifications)

## 2. Session End & Auto-Backup Protocol
After every debug session, bug fix, feature addition, or before concluding a session:
1. Verify:
   - Run `dotnet test engine/zOfit.Engine.Tests/zOfit.Engine.Tests.csproj` (must pass 100%).
   - If release artifacts changed, run `build/build_release.ps1`.
2. Document:
   - Update `memory.md` (Section 2 Snapshot & Section 6 Session Log).
   - Update `tasks/todo.md`.
3. Auto-Backup to Git:
   - Run:
     ```powershell
     git status
     git add .
     git commit -m "<type>(<scope>): <clear description of changes>"
     git push origin main
     ```
   - Report the commit message and push status to the user.

## 3. Strict Project Constraints
- **Excel Output**: All generated formulas must be 100% native live Excel formulas (no dead numbers, no `LET`/`IFS`/`XLOOKUP`, no UDFs).
- **VBA Code**: Pure ASCII only (byte values $\le 127$).
- **Git File Limits**: Never commit files $> 100$ MB (`dist/setup/`, `dist/engine/`, `bin/`, `obj/`).
- **Offline / Privacy First**: AutoCAD MCP parser runs 100% locally via C# on port 9877. Zero cloud dependency.
