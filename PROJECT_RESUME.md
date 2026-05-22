# Generator App — Project Resume Doc

> Single source of truth to pick the project back up after a weekly-limit reset.
> Last updated: 2026-05-22

---

## 1. What this app is

A **single-file React web app** (`index.html`) that generates Google-Sheets-compatible `.xlsx` reports for a 24/7 network monitoring team. It runs entirely client-side — open `index.html` in a browser and you have the full UI.

There are now **two modules**, switched via a top-right toggle:

| Module     | Purpose                                          | File output                              |
|------------|--------------------------------------------------|------------------------------------------|
| SIP / FCS  | Monthly hourly-block / flat-table report builder | `{Month}_{Year}_SIP_FCS_Hourly_Record.xlsx` |
| BMR        | Daily Balance Day & Night blank-data generator   | `BMR_{YYYY-MM-DD}.xlsx`                  |

---

## 2. Tech stack (already wired)

- **React 18** (UMD build) + **Babel standalone** for inline JSX
- **Tailwind** (CDN, no build step)
- **exceljs 4.4.0** (CDN) — used for both modules' xlsx generation
- **file-saver 2.0.5** for downloads
- **Firebase 10.12.5 compat builds** — auth + Firestore for cross-device sync (lazy-loaded; runs fine without)
- **localStorage** key `mrg_state_v1` — primary persistence (works offline)

No build step. To run: just open `index.html`. To deploy: serve the file statically.

---

## 3. File map

```
Generator App/
├── index.html                       ← THE WHOLE APP (one file, ~2900 lines)
├── monthly-report-generator-spec.md ← Original SIP/FCS-only spec
├── letter-g.png                     ← Favicon source (unused at runtime — favicon is inline SVG)
└── PROJECT_RESUME.md                ← This file
```

The reference Excel for the BMR module lives at:
`C:\Users\paxif\Downloads\Balance Day & Night DONT EDIT DELETE.xlsx`

---

## 4. SIP / FCS module (existing — don't touch unless asked)

Already documented in `monthly-report-generator-spec.md`. Key entry points in `index.html`:

- `DEFAULT_SHEETS` (~line 282) — 11 pre-configured sheets (SBC, VOS, ACTIVE S, Feb ALL OUT, VC SIPFCS, VOS2 GSMGW, VOS CYN GSMGW, KINGSFORD INOUT, UNOBANK IN, MYVELOX INOUT, SBC ALARM).
- `generateWorkbook(state)` — top-level xlsx builder.
- `buildHourlySheet` / `buildFlatSheet` / `buildAlarmSheet` — per-layout builders.
- `function App()` (search for it) — top-level component; renders SIP/FCS view when `state.module !== 'bmr'`.

State shape (SIP/FCS):
```js
{
  year, month,
  sheets: [...],            // DEFAULT_SHEETS-shaped
  selectedSheetId,
  rules: [...],             // formatting rules
  includeIndex, changelog, tab, theme, imported,
  module: 'sip_fcs' | 'bmr',
  bmr: { ... },             // BMR-only slice
  bmrTab: 'clients' | 'rules' | 'preview' | 'notes',
}
```

---

## 5. BMR module (new — added 2026-05-22)

### 5.1 What the user wanted

> *"new module called BMR. Generate blank data daily. Manage clients in a separate tab (color, allowable balance). Highlight matching client text. Conditional formatting on Amnt Receivable columns (>= N → color). 30-min usage range rules. Insert new client between specific rows. Drag and drop. Hide rows/columns."*

Source spreadsheet (`Balance Day & Night DONT EDIT DELETE.xlsx`) has:
- 2 sheets: **7AM-7PM** (Day) and **7PM-7AM** (Night)
- Row 1: time labels every 30 minutes (25 slots per shift)
- Row 2: per-block headers
- Rows 3+: one row per carrier
- 318 columns total = 2 leading + 25 timeslot blocks × 12 cols each (+ trailing padding)

### 5.2 Column layout — memorize this

Each time-slot block is **12 columns wide** starting at **col C** (col 3):

| Offset | Header             | Day-slot-1 col |
|--------|--------------------|----------------|
| 0      | CarrierName        | C              |
| 1      | Credit Limit       | D              |
| 2      | Critical Balance   | E              |
| 3      | Billing Cycle      | F              |
| 4      | IsPrepay           | G              |
| 5      | Inv Upto           | H              |
| 6      | Prev Balance       | I              |
| 7      | Pay&Adj            | J              |
| 8      | Current Usage      | K              |
| **9**  | **Amnt Receivable**| **L**          |
| **10** | **30mins usage**   | **M**          |
| 11     | Total usage        | N              |

Cols A and B sit *outside* the blocks:
- **A**: `CarrierName ` (master row label)
- **B**: `Allow Bal` (per-client text label like "Block once debit")

#### Target columns (Amnt Receivable) — 25 cols, +12 stride
`L, X, AJ, AV, BH, BT, CF, CR, DD, DP, EB, EN, EZ, FL, FX, GJ, GV, HH, HT, IF, IR, JD, JP, KB, KN`

#### 30-min usage columns — 25 cols, +12 stride
`M, Y, AK, AW, BI, BU, CG, CS, DE, DQ, EC, EO, FA, FM, FY, GK, GW, HI, HU, IG, IS, JE, JQ, KC, KO`

These are computed at runtime in `index.html` as:
```js
const BMR_TARGET_COL_LETTERS = Array.from({length: 25}, (_, k) => colLetter(12 + 12 * k));
const BMR_USAGE_COL_LETTERS  = Array.from({length: 25}, (_, k) => colLetter(13 + 12 * k));
```

### 5.3 Auto-formulas inside the generated file

Mirrors the source file's pattern:
- **30mins usage** (cols M, Y, AK, …): `=<prevAmnt><row>-<curAmnt><row>` (e.g. `Y3 = L3-X3`)
- **Total usage** (cols N, Z, AL, …): `=<prevTotal><row>+<curUsage><row>` (e.g. `Z3 = N3+Y3`)

First slot has no "previous", so cells there are left blank.

### 5.4 State shape (`state.bmr`)

```js
{
  clients: [
    { id, name, color, allowBalance, allowBalanceLabel, hidden },
    ...
  ],
  targetRules: [   // each rule targets one client across all 25 Amnt Receivable columns
    { id, name, clientId, kind: 'gte'|'gt'|'lte'|'lt'|'eq'|'between',
      value, min, max,          // value for single-bound, min+max for between
      color, fontColor,         // hex like '#EF4444'
      bold, italic, underline,
      enabled },
    ...
  ],
  usageRules: [    // applied to all 25 "30mins usage" columns — same numeric/style fields without clientId
    ...
  ],
  hiddenCols: ['L','X', ...],   // letters — hidden in generated file
  includeDay: true,
  includeNight: true,
  notes: '',
}
```

Default state lives in `DEFAULT_BMR_STATE` (search for it in `index.html`).
Pre-seeded with one COMMSOL-DID target rule (`>= 5` → red bold) and one usage rule (between `-49 and -0.001` → amber bold) to demonstrate.

### 5.5 Key functions (search by name in `index.html`)

| Function                       | What it does                                                                 |
|--------------------------------|------------------------------------------------------------------------------|
| `bmrSlotsForShift(shift)`      | Builds the 25 time-slot strings for 'day' (`7AM` start) or 'night' (`7PM`)   |
| `argbHex(hex)`                 | Converts `#RRGGBB` → `FFRRGGBB` for exceljs                                  |
| `bmrBuildConditionFormula(rule, cellRef)` | Compiles a rule → `AND(ISNUMBER(...), ... )` formula              |
| `bmrStyleFromRule(rule)`       | Compiles a rule → exceljs `{ fill, font }` style object                      |
| `bmrEvaluateRule(rule, value)` | JS-side evaluator (used by preview, not generation)                          |
| `buildBmrSheet(ws, bmr, shift)`| Writes headers, client rows, formulas, conditional formatting on one sheet  |
| `generateBmrWorkbook(state)`   | Top-level: builds Day + Night sheets, downloads `BMR_YYYY-MM-DD.xlsx`        |

### 5.6 UI structure

When `state.module === 'bmr'`, App renders:
- **BmrSidebar** (left, fixed 320px)
  - Brand, sync badge, theme toggle
  - Today's date
  - Shift toggles (Day / Night)
  - Hidden-columns chip list + add input
  - Quick clients list (read-only summary)
  - Generate button
- **Main panel** with 4 tabs:
  - `clients` → `BmrClientsEditor` — drag-and-drop list, hide toggle, color picker per row, **insert at row #** input
  - `rules`   → `BmrRulesPanel` — two sections: Target column rules + 30-min usage rules
  - `preview` → `BmrPreview` — mini HTML table of first 8 clients × first 3 timeslots, with client colors applied
  - `notes`   → free-form textarea bound to `state.bmr.notes`

Switching between SIP/FCS and BMR is via the chip group in the top-right of either header.

### 5.7 Conditional-formatting strategy

For each enabled target rule with a selected client, the generator loops the 25 target columns and adds one `addConditionalFormatting` call per column with a formula like:
```
AND(AND(ISNUMBER(L3), L3 >= 5), EXACT($A3, "COMMSOL-DID"))
```
The data range is `L3:L{N+2}` where N = visible client count, and the column-A name test scopes it to that client. Usage rules remain column-wide across the 25 usage columns.

Cell text-match coloring for clients is *both* applied as a static fill on column A *and* added as an `EXACT(A3, "CLIENTNAME")` conditional formatting rule on the A column range — so the color survives if a user edits the cell.

All formulas use `AND`, `ISNUMBER`, `EXACT` — all of which work in both Excel and Google Sheets.

### 5.8 Hide rows / columns

- **Hide row**: each client has a per-row `hidden: true` toggle (eye icon). Hidden clients are dropped at generation time.
- **Hide column**: sidebar input accepts letters like `L, X, AJ`. They're stored in `state.bmr.hiddenCols` and applied with `ws.getColumn(letter).hidden = true`.

### 5.9 Insert client at row #

In the Clients tab footer:
- Free-text "New client name"
- Optional "Insert at row #" (1-based). Empty = append.
- Hitting Insert (or Enter on the name field) inserts the new client at the chosen index.

### 5.10 Drag-and-drop

Uses the same vanilla HTML5 DnD pattern as `Sidebar`/`ListEditor` in the existing code — grab handle activates the row, dragover highlights the drop target with a blue border. No external library.

---

## 6. How to resume

If the user comes back after a quota reset and says "continue the BMR module," likely next asks:
1. **More rules** (e.g. text contains, not just numeric ranges). Add a new `kind` to `bmrBuildConditionFormula` and handle it in `bmrStyleFromRule` if it needs anything more than fill/font.
2. **Time-of-day rules** (only color certain timeslots). The 25-column loop already isolates by column, so just filter `BMR_TARGET_COL_LETTERS` by index.
3. **Templates** for BMR (mirror the SIP/FCS `TemplateBar`). Storage key `mrg_bmr_templates_v1` is currently unused.
4. **Import from existing BMR xlsx** — re-use `parseXlsxToSheets` pattern: detect time-label row 1 and 12-col stride; pull A-col carriers into `bmr.clients`.

---

## 7. Quick verification recipe

After editing, sanity-check from a terminal:

```bash
# Syntax-check the JSX
node -e "
const fs=require('fs');
const html=fs.readFileSync(String.raw\`C:\\Users\\paxif\\Downloads\\Pao Claude Project\\Generator App\\index.html\`,'utf8');
const m=html.match(/<script type=\"text\/babel\"[^>]*>([\s\S]*?)<\/script>/);
require('@babel/parser').parse(m[1],{sourceType:'module',plugins:['jsx']});
console.log('OK');
"

# Open in browser
start "" "C:\Users\paxif\Downloads\Pao Claude Project\Generator App\index.html"
```

Live smoke test:
1. Click **BMR** in the top-right chip group → BMR view loads.
2. **Clients** tab → drag a row to reorder; toggle hide; change a color; insert at row 2.
3. **Rules** tab → flip enabled, change `between` min/max, change colors.
4. **Preview** tab → see your changes reflected.
5. **Generate** in the sidebar → downloads `BMR_<today>.xlsx`. Open in Excel and Google Sheets — colors should appear when cells are filled with values that trigger the rules.

---

## 8. Known limits / non-goals

- BMR module **does not** populate the timeslot blocks with real data — it generates a blank scaffold (with formulas in usage/total columns). The user enters readings manually.
- **No** import of existing BMR xlsx into the editor yet (planned in §6).
- The 30-min usage formula uses the *prior column's Amnt Receivable* — this matches the source file. If the user later wants a different basis, edit `buildBmrSheet`.

---

## 9. Storage / sync notes

- localStorage: BMR state is folded into the same `mrg_state_v1` object under key `bmr`.
- Firestore: same — the whole state doc syncs after sign-in.
- Old saves (pre-BMR) hydrate fine — `DEFAULT_BMR_STATE` is merged in if `saved.bmr` is missing.

---

## 10. One-line summary

> Single-file React+Tailwind+exceljs app. Two modules: SIP/FCS (monthly hourly/flat reports) and BMR (daily Day+Night balance scaffold with 25 time-slot blocks, per-client colors, client-scoped Amnt Receivable formatting, and column-wide 30-min usage formatting). All client-side. Toggle modules from the top-right chip in the header.
