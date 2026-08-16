---
name: verifying-against-powerpoint
description: Use when changing what paper-pptx accepts, refuses, or writes — any guard, validation, refusal, or save-path change, and any plan in reference/paper-pptx-audit/planned/. Also use when an argument rests on "upstream python-pptx opens it" or "stdlib zipfile accepts it", when deciding whether a refusal is over-strict, and when judging whether a package shape is safe.
---

# Verifying against PowerPoint

## Overview

paper-pptx exists to stop agents producing corrupt decks. So the question is never "does
this file open in Python?" — it is **"does this file open in PowerPoint?"** Those answers
differ, and the gap is the whole reason the package exists.

**PowerPoint is the only oracle. Every other reader is evidence, not proof.**

## The compatibility contract

Two rules. The second overrides the first.

1. **If upstream python-pptx opens it, paper-pptx must open it.** Refusing a file upstream
   reads is a regression against users.
2. **Unless PowerPoint refuses it or reads it wrong.** Then paper-pptx must refuse it too,
   with an error that tells the caller how to fix it. Matching upstream into a corrupt file
   is worse than refusing.

Rule 1 alone is the trap. It is defensible, it is what `CONVENTIONS §1.1` says, and applied
without rule 2 it will make you delete guards that are protecting users.

## Oracle hierarchy

| Reader | Authority | What a pass means |
|---|---|---|
| **PowerPoint** | decisive | the file is good |
| paper-pptx / python-pptx | contract check | tells you nothing about validity |
| LibreOffice | weak, permissive | renders ⇒ *maybe* fine. fails ⇒ definitely broken |
| stdlib `zipfile` | weakest | it parsed. nothing more |

LibreOffice's real use is comparing *content*: two decks that both render with identical
extracted text were read the same way. It cannot tell you PowerPoint will accept either.

## When to use

- Changing any guard, refusal, or validation in the load or save path
- Working any plan in `reference/paper-pptx-audit/planned/`
- Any reasoning step of the form "upstream opens it, so we should"
- Deciding whether an existing refusal is over-strict
- Adding a refusal (does PowerPoint agree the file is bad?)

**Skip only when** the change cannot alter which byte sequences are accepted or produced —
docstrings, type hints, test-only refactors.

## Procedure

**1. Check the ledger first.** `POWERPOINT-VERDICTS.md` records every shape already put in
front of PowerPoint. If your case is there, use it. If it is not there, it is unverified.

**2. Generate variants.** One mutation per file, all from one control, built with *upstream*
python-pptx so the package under test never shaped the input. `--list` shows the mutations;
add one by registering a function in `MUTATIONS`. Keep each minimal — a variant that changes
two things isolates neither.

**3. Run the reader matrix**, then **4. render check**. Both distributions own the `pptx`
import name, so steps 3a/3b need separate environments.

```bash
D=~/Desktop/verify-<change>
uv run --no-project --with 'python-pptx==1.0.2' scripts/make_variants.py $D directory_entries orphan_part
uv run --no-project --with 'python-pptx==1.0.2' scripts/probe_readers.py $D   # upstream
uv run --project . scripts/probe_readers.py $D                                # under test
scripts/render_check.sh $D
```

Compare the **text hash**, not just open/refuse. A reader that opens an ambiguous archive
and shows the wrong content looks like success and is the exact failure to prevent.

**5. Hand the folder to the human.** The generator writes a `README.txt` with per-file
expectations and the three verdicts to report: `OPENS`, `REPAIR`, `OPENS-WRONG`. **Stop here
and ask.** Do not write a recommendation while this step is outstanding.

**6. Record the verdicts** in `POWERPOINT-VERDICTS.md`, then decide.

## Error messages are part of the deliverable

A refusal exists so an agent can repair the file without a human. Every refusal must name
what was found, why it is unsafe, and what to do. `"ZIP member name 'ppt/' is noncanonical"`
fails all three — there is no documented notion of canonical to conform to, and no next
step. Prefer: what is wrong → why it cannot be interpreted → the fix.

Same standard for the exception type: `PackageLimitError` means "too large or ambiguous to
expand". Using it for a structural defect sends the caller to retry with a smaller file.

## Common mistakes

| Mistake | Why it bites |
|---|---|
| "stdlib/LibreOffice accepts it, so it's valid" | Both are permissive. Neither predicts PowerPoint. |
| "No real producer emits this shape" | Corruption and post-hoc byte edits do. Producer surveys measure authoring, not what arrives. |
| Surveying thousands of files instead of opening one | Breadth is not the same measurement. A corpus scan cannot answer "will PowerPoint accept it". |
| Writing the recommendation before the human replies | The verdict can invert it. Wait. |
| Multiple mutations in one variant | You learn that *something* broke, not what. |
| Deleting a guard because upstream is permissive | Upstream is permissive about things PowerPoint rejects. That gap is the product. |
| Building the control with paper-pptx | The package under test must not have shaped the input. |

## Red flags — stop and generate a verification set

- "This can't matter" / "provably undamaged" / "cosmetic"
- "Upstream opens it, so refusing it is unauthorized"
- "No spec authorizes this guard"
- A table of readers with no PowerPoint column
- A recommendation to delete a refusal, of any size

All of these mean: build the decks, ask the human, then decide.

## Real-world impact

An audit concluded that paper-pptx's end-of-central-directory guard refused intact files
and should go. Evidence: 52,807 real files surveyed, five genuinely-intact SEC filings
refused, every member CRC-clean, byte-identical part hashes, stdlib and upstream both
opening the files, and LibreOffice rendering them to identical text. A fix shipped as a PR.

Then PowerPoint refused all four mutated decks and offered only a Repair dialog. The guard
had independently discovered PowerPoint's real rule, which no Python reader reproduces. The
PR was wrong, and every piece of evidence behind it was true.

Breadth of evidence does not substitute for the one measurement that decides the question.
