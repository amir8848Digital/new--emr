# EMR — VB6 → modern web migration

> `CLAUDE.md` in this directory is a symlink to this file — one source of truth, and Claude Code
> auto-loads it at session start. Edit `AGENT.md`; never replace the symlink with a copy.

`/Users/amir/emr` is a **container directory, not a repo.** It holds two independent git repos plus a
read-only legacy source dump. Always confirm which one a task belongs to before editing.

| Directory | What it is | Status |
|---|---|---|
| `emr_node/` | **The backend.** Node/Express + a doctype framework. 235 doctypes. | Active — default target |
| `Basecamp_Frontend/` | **The frontend.** Next.js 15 App Router, React 19, JSON-driven forms. | Active |
| `61BACK ForRn/` | **The legacy VB6 source.** The original app, verbatim. | Read-only reference — never edit |
| `SriRnBackup/` | **A partial VB6 backup snapshot.** A cut-down "Sri" project + the MWF custom-control library. | Read-only reference — never edit |

Remotes are `github.com/8848digital/{emr_node,Basecamp_Frontend}`. **Don't record current branch
names anywhere** — they change under you mid-session (measured: `emr_node` moved `fix/master-p` →
`ai_implementation` inside 15 minutes on 2026-07-29). Run `git branch --show-current` when you need it.

## The project

Porting a legacy VB6 EMR/jewellery-manufacturing system to web. Each VB6 form becomes a **doctype**:
the backend emits a JSON form config + business logic; the frontend renders it dynamically. There are
no hand-written React forms for doctypes.

## `61BACK ForRn/` — the legacy code

The complete original VB6 project (~907 files): 219 `.frm` forms + their `.frx` binaries, `.vbp`
project files (`EmrFullPrj.vbp` is the main one), `.bas` modules, `.cls` classes (`EmrClsHlp.cls` is
the F2-help engine), and `.dsr`/`.dsx` DataReport designers.

- **Never edit anything in here.** It is not a git repo — there is no undo.
- Note the space in the name: quote the path (`"61BACK ForRn/…"`) in every shell command.
- Expect duplicates and dated snapshots (`EmrClsHlp-310118.cls`, `EmrFrmCertBag-old.frm`,
  `-latest`, `Copy of …`). Prefer the plain, undated name unless a doc says otherwise —
  **a name or date suggesting "newer" usually means a smaller SUBSET here.** Measured:
  `EmrClsHlp-latest.cls` (2018) has 101 help routines against `EmrClsHlp.cls`'s 107 (2023).
  Compare content — count routines / `Set wfrm = New` cases — never the filename or mtime.
- Use it to answer "what did the original actually do?" for a form that has **not** been migrated,
  or when a doctype's own `reference/` copy is missing or looks truncated.

For a doctype that is already in flight, prefer its checked-in copy instead — see below.

**A curated ~2.4 MB slice is vendored inside the backend at `emr_node/reference/legacy/`** —
the MWF control library (`ctl/`, irreplaceable) plus the files every migration opens regardless
of form (`core/`: both menu trees, both dispatchers, `Mod.bas`, `EmrClsHlp.cls`). Use those for
menu/dispatch/help lookups; come back out to `61BACK ForRn/` for any individual form.
See `emr_node/reference/legacy/README.md`.

## `SriRnBackup/` — a partial VB6 backup

173 files, not a git repo, no undo. **Never edit anything in here.** Note the casing: `SriRnBackup`,
not `sriRnBackup`. It is a nested backup of a *subset* of the VB6 app, not a second full dump:

| Path | What it holds |
|---|---|
| `61BACK/` + `61CtlBack/` at the top level | one snapshot |
| `SriRnBack/61BACK/` + `SriRnBack/61CtlBack/` | a **second snapshot with newer timestamps but, for the shared files, LESS content** (see below), plus `61RnBack.zip` |

- `61BACK/` is the short project `EmrSriShortPrj.vbp` — only ~6 forms (`EmrFrmDsgMst`,
  `EmrFrmTrackStatus`, `EmrSuppFrmUpd`, `EmrFrpWipBPcs`, `emrfrmmdi`, plus
  `Emr 521 Merge Copy EmrFrmOrdMst.frm`) with the usual `.bas`/`.cls` support files. Everything here
  also exists in `61BACK ForRn/`, so **`61BACK ForRn/` stays the default reference.**
- `61CtlBack/` is the one thing that is **genuinely absent from `61BACK ForRn/`**: the MWF custom
  control library the forms are built on — `MWCTL_*.ctl` (`FLX`/`FLX1` grids, `HLP` F2-help, `BTN1`,
  `ZOOM`, `ENM`, `REP`, …), `MWCLS_*.cls`, and the `MDO*` data-access classes. Go here when you need
  to know how a legacy control actually behaved, not just how a form used it.
- Copies **diverge** — same filename, different content and dates across the two snapshots and
  against `61BACK ForRn/`. Measured 2026-08-10: `EmrFrmDsgMst.frm` is Jun 2024 in `61BACK ForRn/`
  and in `SriRnBackup/61BACK/`, but Feb 2025 (one line longer) in `SriRnBackup/SriRnBack/61BACK/`;
  `EmrSuppFrmUpd.frm` differs too. `SriRnBack/` carries the newer *timestamps* — but newer is
  **not** more complete: its `EmrMod.bas` has **1** dispatch case against `61BACK ForRn/`'s
  **320**, because these folders are the cut-down 5-form `EmrSriShortPrj` project, not the
  full app. Its `FormSelect.bas` is likewise a strict subset (299 cases, adds none).
  **For anything shared — menus, dispatchers, `Mod.bas`, `EmrClsHlp.cls` — use
  `61BACK ForRn/`.** Diff before trusting either copy, and say which one you read.
- `.frm` files read as binary — use `command grep -a` (plain `grep` is `ugrep -I` and silently
  returns nothing).

## Per-doctype reference material

`emr_node/models/doctypes/<Doctype>/reference/` is the **primary source of truth for what a doctype
is supposed to do** — read it before implementing. 116 doctypes have one; between them they hold 78
`.frm` files plus ~428 markdown requirement/status docs (`STATUS.md`, `MENUMAP.md`,
`WORKAROUNDS.md`, `FE_PUNCH_LIST.md`, parsed VB6 dumps).

If a doctype has no `reference/` folder, fall back to `61BACK ForRn/`.

## The backend ↔ frontend contract

The two repos are coupled through one JSON shape. When changing either side, check both:

- Backend emits it: `emr_node/models/doctypes/<Doctype>/config/formConfig.js`
- Frontend types it: `Basecamp_Frontend/interface/formConfigInterface.ts` ← **authoritative schema**
  (the `Field` interface). Read it before inventing a config property.

## Where the deep detail lives

Don't re-derive what these already cover.

| Path | What it holds |
|---|---|
| `emr_node/CLAUDE.md` | Doctype folder anatomy, framework entry points, backend conventions |
| `Basecamp_Frontend/CLAUDE.md` | Render pipeline, `RHFFormFactory` field categories, modal architecture |
| `emr_node/.agent/agent.md` | The WAT (Workflows/Agents/Tools) operating model for the backend |
| `emr_node/.agent/memory/` | Persistent feedback + project decisions. `MEMORY.md` is the index, but it has **known rot** — prefer `ls .agent/memory/` |
| `emr_node/.agent/workflows/` | Migration SOPs (`vb6-migration_v3.md`, `doctype-logic.md`, …) |
| `emr_node/.agent/tools/` | VB6 parsers, migration scripts, docx↔md converters — check here before writing a new script. `parse_mwctl.py` (any `.frm` → JSON: fields, grid columns, decoded types, frame paths) + `mwctl_enums.py` (the MWCTL decoder tables) |
| `emr_node/.claude/commands/` | `/migrate`, `/migrate-v3`, `/parse-vb6`, `/verify-form`, `/generate-doc`, … (user-typed only) |
| `emr_node/.claude/skills/` | `vb6-formconfig-generation` (**.frm + screenshot → formConfig/meta**), `vb6-parity-checklist`, `vb6-ui-parity`, `doctype-child-tables`, `doctype-table-extension`, `doctype-button-types`, `doctype-record-lock`, `doctype-grid-addrow-autofocus`, `addon-menus-reports` |
| `emr_node/reference/legacy/` | Curated ~2.4 MB legacy slice vendored INTO the repo — `ctl/` (MWF control library, irreplaceable) + `core/` (both menu trees, both dispatchers, `Mod.bas`, `EmrClsHlp.cls`). See its `README.md` |
| `Basecamp_Frontend/.claude/skills/` | `rhf-form-button-types` |

Skills and each repo's `CLAUDE.md` are directory-scoped and load automatically when working under
that repo. **`.agent/` does not** — open it explicitly.

## Running the app

`.claude/launch.json` at this level defines **`basecamp-dev`** — the Next.js frontend on port 3001.
Start it with the preview tooling, never with a raw `npm run dev` in Bash. The backend has its own
launch config under `emr_node/.claude/`.

## Working conventions

- **Cite code by symbol name, not line number.** Measured 2026-07-29: all six
  `BaseDocument.js:NNNN` citations in the `RmMst`/`Rm`/`RmAlyCmp` doctypes pointed at unrelated
  code, and one comment cited a method that has never existed. `BaseDocument.js` is ~3450 lines and
  churns. Use `grep -n "symbolName"`. Those six were corrected in `emr_node@1bda1965`; the two
  doctype skills cite symbols throughout.
- **These repos are edited in parallel by other sessions.** Files have changed mid-session. Re-check
  a file before relying on something you read earlier in a long task.
- **Treat in-code claims about "the frontend doesn't support X yet" as unverified.** Several are
  stale (see the `focusedTable` row-selection protocol, which comments still call a pending gap).
- Both repos: Prettier + ESLint + husky. The working tree is CRLF while Prettier is configured for
  LF, so `prettier --check` fails repo-wide — that is pre-existing. **Do not mass-reformat.**
- `emr_node/` root is littered with `_diag_*.txt` / `query_debug.log` scratch output. Ignore it;
  don't treat it as project state, and don't add more.
