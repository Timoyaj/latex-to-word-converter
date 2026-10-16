# Word VBA Macro Suite — Manual Test Checklist

VBA cannot run in CI, so the macros are validated in Microsoft Word against the
fixture documents in this directory. Run through this script before releasing
changes to `macros/EquationAutoNumbering.bas`.

## Setup

1. Open Word, press <kbd>Alt</kbd>+<kbd>F11</kbd>, then **File → Import File…**
   and import `macros/EquationAutoNumbering.bas`.
2. Confirm the module compiles: **Debug → Compile VBAProject** (must be clean).

## Fixtures

| File | Purpose |
| :--- | :--- |
| `macro_smoke.docx` | Raw `$$…$$` / `$…$` text — exercises `ConvertAllLatexInDocument` |
| `renumber_targets.docx` | App-generated `(1)…(3)` equation tables — exercises renumbering & references |
| `word_native_equation_stub.docx` | Reference for Word's native equation SDT structure |

## 1. ConvertAllLatexInDocument (macro_smoke.docx)

1. Open `macro_smoke.docx` and run `ConvertAllLatexInDocument` (Alt+F8).
2. Click **Yes** (convert + auto-number).
3. ☐ The multi-line `$$ … $$` block becomes one native equation in a 1×3
      borderless table numbered **(1)**.
4. ☐ The single-line `$$a^2 + b^2 = c^2$$` becomes equation **(2)**.
5. ☐ Inline `$x^2 + y^2 = z^2$` and `$\alpha + \beta = \gamma$` are native
      inline equations (click into them — Equation Tools ribbon must appear).
6. ☐ The currency sentence still reads **“the price is $5.00 and the other is
      $10.00 today”** — unchanged, no equation.
7. ☐ The escaped-dollar sentence still shows a literal `$5`.
8. ☐ The summary dialog reports 2 display + 2 inline, 0 failed.
9. ☐ Press <kbd>Ctrl</kbd>+<kbd>Z</kbd> ONCE — the entire conversion undoes.

## 2. ConvertSelectionLatex

1. Undo back to raw text. Select just `$$a^2 + b^2 = c^2$$`.
2. Run `ConvertSelectionLatex`.
3. ☐ Selection becomes a centered native equation.

## 3. RenumberAllEquations (renumber_targets.docx)

1. Open `renumber_targets.docx` (labels `(1)`, `(2)`, `(3)`).
2. Delete the first equation table entirely.
3. Run `RenumberAllEquations`.
4. ☐ Remaining tables read **(1)** and **(2)** — no gap.
5. ☐ Labels are right-aligned in Cambria Math 12pt.

## 4. InsertEquationNumber

1. Place the cursor in an empty paragraph of `renumber_targets.docx`.
2. Run `InsertEquationNumber`.
3. ☐ A 1×3 borderless table appears, numbered `(3)`, cursor in the center cell.
4. Run `RenumberAllEquations`.
5. ☐ The new table keeps its place in sequence.

## 5. InsertEquationReference (live fields)

1. Place the cursor in a paragraph, run `InsertEquationReference`, enter `2`.
2. ☐ Text `Eq. (2)` appears.
3. Right-click it → **Toggle Field Codes** — it must be a `REF EqNum…` field,
   not plain text.
4. Delete the FIRST equation table, run `RenumberAllEquations`.
5. ☐ The reference text now shows the renumbered value (updates with F9 /
   renumber), because it is a live field.

## 6. Custom label formats (interop with the Python app)

1. Generate a doc with `python -m core.converter in.md out.docx --eq-format "[Eq. {n}]"`.
2. Open it in Word and run `RenumberAllEquations`.
3. ☐ Labels keep the `[Eq. n]` format — only the number changes.

## 7. Regression guards

- ☐ Word's undo history contains ONE "Convert LaTeX Equations" entry for a run.
- ☐ No error dialog appears when a formula is invalid (it is skipped and
      counted as failed in the summary).
- ☐ Column widths respect the page size (check in landscape: the table still
      fits between margins).
