# PowerPoint verdict ledger

What the real PowerPoint application does with each package shape we have actually put in
front of it. **Check here before generating a verification set** — a shape already recorded
does not need re-testing, and a shape not recorded here has *not* been verified no matter
how confident the reasoning looks.

Verdicts come from a human opening the file in PowerPoint. Nothing else may be written into
the PowerPoint column — not a LibreOffice render, not a stdlib `zipfile` result, not an
inference from the spec.

| mutation | stdlib `zipfile` | upstream python-pptx | **PowerPoint** | verified |
|---|---|---|---|---|
| *(control, unmodified)* | opens | opens | **opens** | 2026-08-15 |
| `declared_comment` — spec-correct archive comment | opens | opens | **opens** | 2026-08-15 |
| `trailing_bytes` — undeclared bytes after the footer | opens | opens | **REPAIR** | 2026-08-15 |
| `trailing_newline` — one undeclared `\n` | opens | opens | **REPAIR** | 2026-08-15 |
| `bogus_comment_length` — declares a comment that is absent | opens | opens | **REPAIR** | 2026-08-15 |
| `concatenated` — two whole packages in one file | opens | opens, shows only deck 2 | **REPAIR** | 2026-08-15 |
| `directory_entries` — zero-byte `ppt/` folder records | opens | opens | **opens** † | 2026-08-15 |
| `prefix_data` — stub before the archive, offsets rebased | opens | opens | **REPAIR** † | 2026-08-15 |
| `stray_signature` — footer signature inside a member | opens | opens | **opens** † | 2026-08-15 |
| `orphan_part` — part no relationship points at | opens | opens, drops it on save | *not yet tested* | — |

† Recorded from the verdict table in `tests/paper/test_zipguard_end_record.py`, written by a
parallel session that ran the round-2 set. Not observed first-hand by the author of this
ledger. **Confirm with whoever opened them before relying on these three rows** — the whole
point of this file is that only a human-observed PowerPoint result counts.

## What the tested rows establish

PowerPoint requires the end-of-central-directory record's declared comment length to
account for every remaining byte of the file. Declared trailing data is fine; undeclared
trailing data is not, and neither is a declared length with nothing behind it.

That is substantially the rule `_zipguard._find_end_record` already enforces. It was
reached independently, and no Python reader reproduces it — stdlib `zipfile`, upstream
python-pptx and LibreOffice all accept every one of the REPAIR rows above.

**Consequence for the audit:** an argument of the form "upstream opens it, so refusing it
is an unauthorized capability→refusal conversion" is not sufficient on its own. Upstream
opens all four REPAIR shapes. A guard that refuses them is protecting users from files
PowerPoint will reject, which is the package's stated purpose.

If the three † rows hold, two of them cut in opposite directions and both need work:

- **`directory_entries`** — PowerPoint opens it, paper-pptx refuses. The fork is too strict;
  the acceptance change in flight is correct.
- **`prefix_data`** — PowerPoint refuses it, paper-pptx opens it. The fork is too *lenient*,
  the only such row here. LibreOffice corroborates: it renders that deck to 4 pages where
  every other variant renders to 2, so readers genuinely disagree about the content.
  Refusing it is a capability→refusal change and needs the `CONVENTIONS §1.1` escalation.

## Rows that still need a human

- `orphan_part` gates deleting the unreachable-parts load refusal. Upstream opens the file
  and then silently drops the orphan on save, so both the open and the round-trip need
  checking: does PowerPoint open it, and does the part survive a PowerPoint re-save?
