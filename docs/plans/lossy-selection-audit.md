# Lossy Selection and Ambiguous Matching in `paper-pptx`

**Audit date:** 2026-08-21
**Repository:** `paper-instruments/paper-pptx`
**Audited worktree:** `/Users/gavin/dev/paper-pptx-gap-spec`
**Audited commit:** `93eae181` (`codex/pptx-doc-hygiene`)
**Implementation stack base:** PR #38 final remote head `a3c1ae6e` (`codex/pptx-doc-hygiene`); the
audited commit above is its ancestor.
**Upstream comparison point:** `paper-base` at `278b47b1`
**Purpose:** Explain each selection or matching behavior that can silently choose, associate, or retain the wrong presentation object, so maintainers can assess its seriousness and decide whether to fix it.

## Executive summary

The audit did **not** find the kind of pixel-window or character-window heuristics previously seen in spreadsheet tooling. There is no Paper-added PPTX code that chooses an object by rules such as “six characters to the right,” “closest object within N pixels,” or fuzzy text distance.

The PPTX package does, however, have a related class of lossy behavior: several APIs reduce an ambiguous set of exact candidates to a single candidate by position. Typical examples are “the first layout with this name,” “the first compatible placeholder,” and “the paragraph currently at this ordinal.” These choices are deterministic, but they are not stable semantic identities. A harmless reorder or the presence of duplicate names can redirect an edit or make a report describe changes that did not occur.

The most serious finding is `BlockAnchor`. An anchor is intended to guard a text edit against stale inspection data, but it stores only the presentation part, a slide-global paragraph ordinal, and a short hash of paragraph text. If two paragraphs have the same text and their shapes are reordered, an anchor created for one paragraph can remain apparently valid and silently edit the other paragraph. This is a wrong-write safety issue, not merely a reporting issue.

The next most important findings are automatic layout selection during slide import and automatic placeholder matching during layout rebinding. Both select the first candidate when multiple exact or category-compatible candidates exist. Those choices can silently change geometry, inherited formatting, placeholder semantics, and slide furniture.

The diff engine has two additional positional matchers. They cannot modify a deck, but they can report false shape changes after z-order changes and cascades of false paragraph edits after a paragraph insertion. Because these reports are used to validate edits and diagnose model performance, misleading reports can conceal the actual failure mode.

None of the runtime behaviors identified here were introduced by the current three-PR stack. Most were introduced in the Paper bootstrap commit `700b5f1f`. One first-match API, `SlideLayouts.get_by_name()`, is inherited from upstream `python-pptx`. PRs #36–#38 do not add selection heuristics; PR #38 documents anchors without currently warning about the silent-misbinding case described here.

### Findings at a glance

| ID | Behavior | Failure mode | Can mutate a deck? | Severity | Provenance | Fix surface |
|---|---|---|---:|---|---|---|
| LS-01 | Text anchors use part + ordinal + 32-bit text hash | Reordered duplicate text can redirect an edit | Yes | Critical | Paper bootstrap | Medium-wide |
| LS-02 | Import layout inference takes first name/type match | Ambiguous layouts silently select by deck order | Yes | High | Paper bootstrap | Localized to medium |
| LS-03 | Placeholder rebinding takes first compatible slot | Content binds to an arbitrary compatible placeholder | Yes | High | Paper bootstrap | Medium |
| LS-04 | Shape diff fallback matches by kind ordinal | Reordering duplicate/unnamed shapes creates false changes | No | Medium | Paper bootstrap | Medium |
| LS-05 | Paragraph diff matches by paragraph ordinal | Insertions create cascades of false replacements | No | Medium | Paper bootstrap | Medium-wide |
| LS-06 | Named section lookup returns first match | Import can enroll a slide into the wrong duplicate-named section | Yes | Medium | Paper bootstrap | Localized |
| LS-07 | Footer normalization keeps first and deletes duplicates | Valid duplicate furniture is discarded by policy | Yes | Low to medium | Paper bootstrap | Localized to medium |
| LS-08 | `SlideLayouts.get_by_name()` returns first match | Direct callers can receive the wrong duplicate-named layout | Yes, through later caller actions | Low to medium | Upstream `python-pptx` | Localized if additive |

## What “lossy selection” means in this report

A selection operation is lossy when it starts with more information or more candidates than it preserves, then silently commits to one interpretation without enough stable identity to justify that choice.

This report distinguishes four cases:

1. **Silent ambiguous selection.** Multiple candidates satisfy the stated rule, but the implementation returns the first one.
2. **Positional identity.** An object is treated as the same object because it occupies the same ordinal position, even though order can change independently of semantic identity.
3. **Deliberate canonicalization.** The API intentionally reduces multiple objects to one. This is deterministic and documented in code, but it is still destructive and should be assessed as a product policy.
4. **Conservative refusal.** The API detects ambiguity and raises an error or leaves items unmatched. This is the safe pattern and is not considered lossy.

Exact string comparison is not automatically safe. A lookup for an exact name is still ambiguous if names are not unique. Likewise, deterministic behavior is not automatically correct: “always use the first candidate” is repeatable but can still bind to the wrong object.

## Scope and method

The audit covered selection, search, inspection, matching, rebinding, and diff behavior in `src/pptx`, with particular attention to code added after `paper-base`.

The analysis included:

- comparing Paper-added lines against the `paper-base` tag;
- tracing suspicious lines to their introducing commits with Git blame and history;
- reviewing public and internal selectors for layouts, placeholders, shapes, text blocks, sections, and footer furniture;
- searching Paper-added code for geometric, fuzzy, pixel, distance, proximity, and character-window heuristics;
- executing minimal reproductions for anchor misbinding, layout ambiguity, placeholder ambiguity, shape-diff reordering, and paragraph-diff insertion;
- reviewing the current PR stack to determine whether any open PR introduced or worsened the runtime behavior.

### Provenance terminology

- **Upstream** means behavior already present at `paper-base`, which points to upstream `python-pptx` commit `278b47b1dedd5b46ee84c286e77cdfb0bf4594be`.
- **Paper bootstrap** means behavior introduced by Paper commit `700b5f1f`, “Bootstrap paper-pptx as an agent-first PPTX structure editor.”
- **Current PR stack** means PRs #36, #37, and #38, culminating in audited commit `93eae181`.

### Fix-surface scale

The “lift” in this report is not an estimate of time. It describes how much of the codebase and public contract a correct fix can affect.

- **Localized:** one production module plus focused tests and documentation.
- **Medium:** one shared algorithm or helper with multiple callers, plus several test suites and documentation surfaces.
- **Medium-wide:** multiple production modules, serialized/public schemas, compatibility handling, fixtures or golden reports, and downstream consumers.
- **Broad:** a pervasive architectural change across many unrelated APIs. No finding in this audit requires a broad rewrite if addressed conservatively.

## LS-01 — `BlockAnchor` can silently redirect a text edit

### Severity and category

**Severity:** Critical
**Category:** Positional identity plus insufficient validation
**Mutation risk:** Direct silent wrong write

### Where it occurs

- `src/pptx/inspect.py`: `content_hash()`, `BlockAnchor`, `TextBlock`
- `src/pptx/edit.py`: `replace_text_at()` and anchor resolution
- Related recovery behavior: `refind()`

### The lossy behavior

Inspection exposes richer information about a paragraph than the anchor retains. A `TextBlock` knows the owning shape ID, shape name, container type, and container details. The corresponding `BlockAnchor` stores only:

- the OPC part containing the text;
- a slide-global `block_index`;
- an eight-hex-character SHA-256 prefix derived from normalized paragraph text.

Eight hexadecimal characters provide only 32 bits of hash space. More importantly, two semantically distinct paragraphs with identical text intentionally receive the same hash. The paragraph ordinal is then the only discriminator between them.

When the shape tree or paragraph order changes, the object at a given ordinal can change. If the new object at that ordinal has the same text, the hash check succeeds and the edit proceeds against the wrong paragraph.

### What is bad about it

The API presents the hash as an optimistic-concurrency guard, but the guard validates content, not identity. It answers “does some paragraph at this position still contain this text?” rather than “is this the paragraph the caller inspected?”

This is especially dangerous because the failure is silent:

- no stale-anchor exception is raised;
- the requested replacement succeeds;
- the edited paragraph contains plausible output;
- the intended paragraph remains unchanged;
- a later visual check may notice only that the expected change is missing, not that another object was corrupted.

The short hash adds avoidable accidental-collision risk, but increasing the hash length alone would not solve the primary problem. Identical text in different shapes must remain identical under any content-only hash.

The anchor also intentionally excludes volatile field display text from identity. That can be useful, but field type and field position are not included as stable structural identity either. Consequently, text containing fields has an additional route to looking equivalent when its semantic structure differs.

### Negative example

Consider two text boxes:

- shape A contains `same text` and is paragraph block 0;
- shape B contains `same text` and is paragraph block 1.

An agent inspects shape A and stores its anchor. Another edit changes z-order, so shape B is now traversed first. The visible paragraph text is unchanged in both shapes.

The old anchor still points to block 0, and block 0 still hashes to the same value. Calling `replace_text_at(anchor, "changed text")` edits shape B. Shape A—the intended target—remains unchanged.

The reproduction observed exactly this state:

```text
before        [('A', 0, 'same text'), ('B', 1, 'same text')]
after reorder [('B', 0, 'same text'), ('A', 1, 'same text')]
after edit    [('B', 'changed text'), ('A', 'same text')]
```

This is not a theoretical SHA collision. It uses ordinary duplicate text and an ordinary shape reorder.

### Impact

- Edits can be applied to the wrong text box, table cell, group descendant, or notes paragraph.
- Duplicate labels such as `Title`, `Total`, `Q3`, or repeated footer text make the condition realistic.
- Reordering shapes, inserting a paragraph, or changing group traversal can invalidate ordinals without changing text.
- Automated verification may see that the requested text exists somewhere and miss the collateral edit.
- Stored anchors are unsafe across even modest structural edits unless the caller manually refinds and revalidates them.

### Existing mitigation and why it is insufficient

`refind()` is conservative once a caller knows an anchor is stale: it searches for the exact content hash within the same part and refuses zero or multiple matches. That is a good ambiguity policy.

It does not protect this case because the old anchor does not appear stale. The paragraph at its old ordinal has a matching hash, so normal resolution succeeds before `refind()` is relevant.

### Provenance

This anchor implementation was introduced by Paper bootstrap commit `700b5f1f`. It is not inherited from upstream `python-pptx` and was not introduced by PRs #36–#38.

PR #38 changes anchor-related documentation and docstrings but does not change runtime selection. Its relationship to this issue is therefore documentation risk: stronger anchor-safety language should not imply stable object identity until this failure mode is fixed or explicitly disclosed.

### Recommended fix

Introduce a versioned structural anchor and resolve identity before validating content.

A practical v2 anchor should retain container-specific identity:

- **Top-level shape text:** slide part, stable shape ID, and paragraph ordinal within that shape.
- **Grouped shape text:** slide part, a lineage of stable shape IDs from the top-level group to the descendant shape, and paragraph ordinal within the leaf shape.
- **Table text:** slide part, table-frame shape ID, row index, column index, and paragraph ordinal within the cell.
- **Notes text:** notes part plus the owning notes shape or placeholder identity and paragraph ordinal within that container.

After resolving that structural identity, validate a stronger content fingerprint. The fingerprint should:

- use at least 128 bits, or the full SHA-256 digest;
- normalize text consistently;
- include field types and field positions;
- continue excluding volatile displayed field values where that is necessary for stability.

Legacy anchors should be handled conservatively. A v1 anchor should not silently select among duplicate-text candidates. If stable identity cannot be reconstructed uniquely, the edit should raise a typed ambiguity or stale-anchor error and require a new inspection.

Required regression coverage should include:

- the two-identical-text-box reorder shown above;
- paragraph insertion before an anchored paragraph;
- table cells with identical text;
- nested groups containing identical text;
- notes placeholders with identical text;
- fields with changing display values;
- serialization and deserialization of v1 and v2 anchors;
- explicit refusal for ambiguous legacy anchors.

### Fix surface

**Medium-wide—the largest fix surface in this audit.**

The core logic is concentrated in `inspect.py` and `edit.py`, but the anchor is a public/serialized contract. A complete fix affects:

- anchor dataclasses or schemas;
- inspection output and JSON serialization;
- edit resolution;
- legacy-anchor compatibility;
- docs and examples;
- fixtures and golden outputs that contain anchors;
- any downstream consumer persisting anchors.

This does not require changing the entire package. The risk comes from schema compatibility and the number of container types that must be represented correctly.

## LS-02 — Automatic import layout binding silently chooses the first match

### Severity and category

**Severity:** High
**Category:** Silent ambiguous selection
**Mutation risk:** Imported slide can acquire the wrong layout semantics

### Where it occurs

- `src/pptx/compose.py`: `_resolve_layout_binding()`

### The lossy behavior

When the caller does not provide an explicit target layout, slide import tries progressively weaker rules:

1. exact layout name;
2. matching layout type;
3. in bake mode, a blank layout;
4. in bake mode, the first destination layout as a final fallback.

At each fallback tier, the implementation returns the first candidate encountered. It does not require that the candidate be unique, and it does not surface the alternatives.

A direct reproduction with two layouts sharing the same name produced:

```text
same-name candidates
  /ppt/slideLayouts/slideLayout1.xml
  /ppt/slideLayouts/slideLayout2.xml
selected
  /ppt/slideLayouts/slideLayout1.xml via name-match
```

Changing master or layout order changes the selected result without changing the requested name.

### What is bad about it

Layout names and layout types are not guaranteed to be unique across all masters in a presentation. A layout is also more than a label: it participates in placeholder inheritance, geometry, theme resolution, and slide furniture.

Choosing the first exact-name layout conflates “exact string match” with “unique identity.” Falling back to the first same-type or first overall layout is even weaker. The algorithm lacks enough information to know that the selected layout is the intended one.

Adding a geometric similarity heuristic would make this worse. A layout should be selected by explicit or unique semantic identity, not by approximate coordinates.

### Negative example

A destination template contains two masters, each with a layout named `Title and Content`. The first belongs to a legacy corporate master; the second belongs to the current branded master. The source slide came from the current brand, but automatic name matching selects the first layout because that master appears first in the package.

The imported slide may then exhibit:

- content positioned according to the legacy placeholder geometry;
- title and body text inheriting old font or color settings;
- the wrong footer/date/slide-number furniture;
- a different background or master-level decoration.

The import technically succeeds, and all text may still be present, making the semantic misbinding easy to miss.

### Impact

- Layout choice becomes sensitive to package order.
- Duplicate names across masters can silently cross brand or theme boundaries.
- Same-type fallback can bind to a layout with materially different placeholder structure.
- Bake-mode fallback to `dest_layouts[0]` can select a nonblank, semantically unrelated layout.
- Follow-on placeholder reconciliation can amplify the initial mistake.

### Provenance

The automatic import resolver was introduced in Paper bootstrap commit `700b5f1f`. It was not introduced by the current PR stack.

There is a related inherited upstream behavior: `SlideLayouts.get_by_name()` returns the first exact-name match. That upstream API is covered separately as LS-08. Paper's import resolver is its own broader first-match algorithm across destination masters and layouts; it is not merely a call through to the inherited method.

### Recommended fix

Make every automatic tier uniqueness-preserving:

1. Gather all exact-name candidates. Select only if there is exactly one.
2. If none exist, gather all exact-type candidates. Select only if there is exactly one.
3. For bake mode, gather blank candidates. Select only if there is exactly one.
4. If a tier has multiple candidates, raise a typed ambiguity error that lists part names, master identities, layout names, and layout types.
5. If no unique automatic candidate exists, require `target_layout=` or an explicit caller-supplied policy.

The resolver should not silently continue from an ambiguous stronger tier to a weaker tier. If two exact-name candidates exist, trying type matching does not resolve the ambiguity; it merely discards useful information.

Tests should cover duplicate names on one master, duplicate names across masters, duplicate types, multiple blank layouts, layout-order changes, and explicit-target success.

### Fix surface

**Localized to medium.**

Most production changes belong in `_resolve_layout_binding()` and its typed error/reporting path. The broader surface includes import-mode tests, import reports, documentation, and callers that currently rely on automatic fallback. The change is behaviorally significant but does not require a new package-wide abstraction.

## LS-03 — Placeholder rebinding silently chooses the first compatible slot

### Severity and category

**Severity:** High
**Category:** Silent ambiguous selection
**Mutation risk:** Content can be bound to the wrong target placeholder

### Where it occurs

- `src/pptx/rebind.py`: `_compute_mapping()`
- Callers include slide layout rebinding and adopt-theme import reconciliation.

### The lossy behavior

Placeholder mapping first seeks an exact placeholder type and index, which is a strong match. When that fails, it builds weaker candidate sets and chooses their first member:

- the first unused placeholder with the same type;
- otherwise, the first unused placeholder in a compatible family.

For a source `OBJECT` placeholder at index 9 and target placeholders `OBJECT@1` and `OBJECT@2`, the resolver selected `OBJECT@1` solely because it appeared first.

### What is bad about it

Placeholder type or family compatibility does not determine intended placement when several compatible slots exist. Placeholder indexes and XML order can be arbitrary from the caller's perspective, and two content placeholders may represent different columns, regions, or semantic roles.

The fallback is deterministic but under-specified. It can silently transform an ambiguous mapping problem into a concrete deck mutation.

The existing rebind report emphasizes resolved run/font shifts. That is useful, but the most important consequence of choosing the wrong placeholder may be geometry or role, not text styling. A report about font resolution does not make an arbitrary placeholder assignment safe.

### Negative example

A target layout contains two object placeholders: index 1 is the left comparison column and index 2 is the right comparison column. A source slide has a content placeholder whose original index does not exist on the target layout.

Both target placeholders have the same type. The mapper selects index 1 because it is first, even if the source content semantically belongs in the right column. The content moves and resizes to the left slot, while a different object may subsequently claim the right slot or remain unbound.

### Impact

- Rebinding can move or resize content into the wrong region.
- Placeholder inheritance can change fonts, bullets, colors, and autofit behavior.
- Adopt-theme import can produce structurally valid but visually incorrect slides.
- Results can depend on placeholder order rather than explicit mapping intent.
- Multi-column and repeated-content layouts are particularly exposed.

### Provenance

The mapping algorithm was introduced in Paper bootstrap commit `700b5f1f`. It is not inherited from upstream and was not added by PRs #36–#38.

### Recommended fix

Preserve the current exact type-plus-index match. Make both fallback tiers ambiguity-aware:

- if exactly one unused same-type candidate exists, use it;
- if more than one exists, raise a typed ambiguity error and require `placeholder_map`;
- if no same-type candidate exists and exactly one compatible-family candidate exists, use it;
- if multiple family candidates exist, refuse and require an explicit mapping.

The error/report should include source type/index and every candidate's type, index, owning layout, and part name. That gives agents enough information to choose explicitly.

Do not replace the first-match rule with nearest-coordinate, overlap, or pixel-distance matching. Those signals are presentation-specific and can be invalidated by intentional layout changes—the exact class of brittle heuristic this audit is intended to avoid.

Optional geometry-shift reporting would improve diagnostics after a valid mapping, but it should not be used to justify an ambiguous automatic mapping.

### Fix surface

**Medium.**

The core change is in one shared helper, but that helper serves both direct layout rebinding and import reconciliation. The change therefore touches:

- mapping and error types in `rebind.py`;
- rebind tests;
- compose/import integration tests;
- reports and documentation that describe automatic mapping;
- callers that need to provide `placeholder_map` after a new ambiguity error.

## LS-04 — Shape diff fallback matches duplicate or unnamed shapes by ordinal

### Severity and category

**Severity:** Medium
**Category:** Positional identity
**Mutation risk:** None; report correctness risk

### Where it occurs

- `src/pptx/diff.py`: `_shape_keys()` and structural diff reporting

### The lossy behavior

The diff engine uses a unique shape name when one is available. When a name is absent or duplicated, it falls back to a key composed of shape kind and ordinal among shapes of that kind.

That ordinal reflects traversal or z-order, not persistent semantic identity. Reordering two same-kind shapes changes their ordinals, so the diff pairs each “before” shape with a different “after” shape.

### What is bad about it

The fallback forces a one-to-one association despite not having enough evidence. It produces confident, detailed facet changes rather than acknowledging that matching is ambiguous.

This is report-only, so it cannot corrupt a deck. It can still corrupt the maintainer's understanding of what changed and can mislead automated evaluation that treats the diff as ground truth.

### Negative example

Two duplicate-named text boxes occupy different horizontal positions. The only edit is reversing their z-order. The diff reported two geometry changes:

```text
shape sp#0: left 914400 -> 4572000
shape sp#1: left 4572000 -> 914400
```

Neither shape moved. The matcher paired the old first shape with the new first shape and the old second shape with the new second shape.

The same problem can associate the wrong image, chart, or other same-kind object after a reorder.

### Impact

- False geometry and formatting changes after z-order edits.
- Real changes can be attributed to the wrong shape.
- Diff-based validation can fail a correct edit or bless the wrong explanation.
- Debugging output becomes most misleading exactly when names are duplicated or absent.

### Provenance

The fallback matcher was introduced in Paper bootstrap commit `700b5f1f`. It was not introduced by upstream `python-pptx` or by the current PR stack.

### Recommended fix

The deck diff already treats itself as a lineage comparison by matching slides through permanent slide identity. Within that same lineage, use stable shape ID as the primary identity.

Recommended behavior:

- match top-level shapes by stable shape ID;
- use names only as human-readable labels, not as identity;
- for grouped shapes, use a shape-ID lineage or another exact structural path;
- where an ID is missing, reused, or cannot be trusted, mark the shapes unpaired or ambiguous rather than pairing by ordinal;
- report additions/removals explicitly instead of manufacturing facet changes.

Do not introduce geometric nearest-neighbor or fuzzy name matching. A conservative “unpaired” result is more useful than a precise-looking false match.

### Fix surface

**Medium.**

The matching logic is concentrated in `diff.py`, but changing identity keys affects report labels, serialized output, golden fixtures, and tests for shapes, charts, and pictures. It should not require changes to mutation APIs.

## LS-05 — Paragraph diff matching by ordinal causes cascading false edits

### Severity and category

**Severity:** Medium
**Category:** Positional identity
**Mutation risk:** None; report correctness risk

### Where it occurs

- `src/pptx/diff.py`: `_text_blocks_by_stable_key()`
- Related alignment behavior exists in rebind reporting.

### The lossy behavior

Text blocks are keyed by owning shape ID and paragraph ordinal within that shape. This is more stable than a slide-global paragraph index, but an insertion or reorder inside the same shape shifts all later ordinals.

The matcher then compares different logical paragraphs as if they were the same paragraph.

### What is bad about it

An ordinal says where a paragraph is now, not which paragraph it is. Using it as identity turns one insertion into a cascade of replacements.

The report appears exact because it lists before and after text for each keyed paragraph. It does not expose that its association rule became invalid after the insertion.

Related rebind alignment combines shape ID, block ordinal, and run identity and skips some duplicate or empty identities. That is more conservative in some cases, but block ordinal still limits its ability to align after paragraph insertions. Bullet-diff logic already demonstrates a better principle by requiring unique text and skipping duplicates rather than forcing matches.

### Negative example

Before:

```text
A
B
```

After inserting `X` before both paragraphs:

```text
X
A
B
```

The ordinal matcher reports:

```text
A -> X
B -> A
None -> B
```

The true semantic change is one insertion: `X`.

### Impact

- A single insertion or deletion creates many false text replacements.
- Real edits later in the text body can be attributed to the wrong paragraph.
- Evaluation reports overstate edit scope and obscure the actual action taken.
- Effective-format or bullet shift analysis may become noisy or incomplete when it depends on related positional alignment.

### Provenance

The paragraph matching behavior was introduced in Paper bootstrap commit `700b5f1f`. It is not an upstream behavior and was not added by the current PRs.

### Recommended fix

Align paragraph sequences within an already stable container rather than treating ordinals as identity.

A conservative exact alignment can use longest-common-subsequence or an equivalent sequence-diff algorithm over structural paragraph fingerprints. The fingerprint should include:

- normalized literal text;
- field types and positions rather than volatile field display values;
- paragraph-level structural features when needed to disambiguate;
- the exact container coordinate, such as table cell, before sequence alignment begins.

Duplicate fingerprints must be handled conservatively. If repeated paragraphs make correspondence ambiguous, leave them unmatched or label the region ambiguous. Do not use fuzzy text similarity to force a pairing.

The same exact sequence-alignment primitive should be shared by text, effective-format, and bullet-shift reporting where their requirements overlap. That keeps the fix DRY and prevents three subtly different notions of paragraph identity.

### Fix surface

**Medium-wide.**

The initial algorithm belongs in `diff.py`, but a DRY implementation may also serve related rebind/effective-resolution reporting. Expected affected surfaces include:

- paragraph and text-block diff logic;
- shared fingerprint/alignment utilities;
- effective-format and bullet reports where alignment is reused;
- report schemas or ambiguity markers;
- golden fixtures and a broad set of diff tests.

The runtime editing model does not need to change for this issue.

## LS-06 — Section lookup by name silently returns the first duplicate

### Severity and category

**Severity:** Medium
**Category:** Silent ambiguous selection
**Mutation risk:** Imported slide can be enrolled into the wrong section

### Where it occurs

- `src/pptx/compose.py`: `_find_section()` and slide import section enrollment

### The lossy behavior

Named section lookup iterates sections and returns the first exact name match. The repository enforces invariants around section slide IDs, but the audit found no contract that makes section names unique.

Therefore, if a presentation contains duplicate section names, the requested name does not uniquely identify a section and package order determines which section is returned.

### What is bad about it

The API accepts a human-readable label as though it were stable identity. Exact matching prevents typos or fuzzy matches, but it does not resolve duplicate labels.

Because slide import then mutates section membership, the result is a silent structural error rather than merely a lookup inconvenience.

### Negative example

A deck contains two sections named `Appendix`: one for financial tables and one for legal material. Importing a new legal slide with `section="Appendix"` enrolls it into the first section, which happens to be the financial appendix.

The slide content remains correct, so automated content checks may pass while the visible deck organization is wrong.

### Impact

- Imported slides can appear in the wrong logical section.
- Behavior changes if sections are reordered.
- Agents cannot diagnose or choose among duplicates because alternatives are not surfaced.
- The error may be visible only in PowerPoint's section organization UI.

### Provenance

The first-match section helper was introduced in Paper bootstrap commit `700b5f1f`. It is not inherited from upstream and was not introduced by PRs #36–#38.

### Recommended fix

Collect all exact matches:

- zero matches: raise the existing not-found error;
- one match: return it;
- multiple matches: raise a typed ambiguity error listing section IDs, positions, and names.

For complete control, allow callers to select by stable section ID or pass a section object in addition to the convenience name API.

### Fix surface

**Localized.**

The change belongs primarily in `compose.py`, with focused import tests and documentation. Adding section-ID selection may touch a public signature and report schema, but it remains a small, coherent surface.

## LS-07 — Footer furniture normalization keeps the first placeholder and deletes the rest

### Severity and category

**Severity:** Low to medium, depending on intended product policy
**Category:** Deliberate destructive canonicalization
**Mutation risk:** Duplicate footer/date/slide-number placeholders are removed

### Where it occurs

- `src/pptx/hf.py`: layout furniture collection and slide furniture application

### The lossy behavior

Footer/date/slide-number handling records the first layout placeholder of each furniture kind. On a slide, it gathers all matching placeholders, edits the first, and deletes extras.

Unlike the prior issues, this behavior is not accidental or flaky in implementation. Comments and tests describe a deterministic one-per-kind “dialog state.” It is nevertheless lossy because valid existing objects and their distinct geometry, styling, or content are discarded.

### What is bad about it

The algorithm assumes that multiple placeholders of a furniture kind represent accidental duplication and that the first one is canonical. OOXML decks in the wild may contain intentional duplicates, template artifacts, or placeholders whose order does not reflect visual importance.

The caller asks to configure footer furniture, but the operation also performs structural cleanup. If that cleanup is not an explicit part of the public contract, the mutation exceeds the narrow intent implied by the call.

### Negative example

A custom template has two footer placeholders:

- the first is an invisible or legacy placeholder retained by the template;
- the second is the visible branded footer in the correct position and style.

Applying footer settings updates the first placeholder and deletes the second. The visible footer disappears or moves to the legacy position even though the caller supplied valid footer text.

### Impact

- Intentional duplicate furniture can be deleted.
- The retained placeholder may have the wrong geometry or style.
- Template-specific structure is normalized without an explicit conflict decision.
- The action is difficult to reverse without the original deck.

### Provenance

This behavior was introduced in Paper bootstrap commit `700b5f1f`. It is not inherited from upstream and was not introduced by the current PR stack.

### Recommended fix

First decide and document the product contract.

If one-per-kind canonicalization is an intentional feature, preserve it only behind an explicit operation or option such as `canonicalize=True`, and report exactly which placeholder IDs will be removed. The default configuration path should not silently delete ambiguous duplicates.

For a safer default:

- accept a stable placeholder index/ID when multiple instances exist;
- if no explicit selection is supplied and multiple candidates exist, raise a typed ambiguity error;
- offer a separate canonicalization method for callers that deliberately want one-per-kind normalization.

Do not decide which duplicate to keep by visual proximity or pixel geometry. Explicit identity or conservative refusal is safer.

### Fix surface

**Localized to medium.**

The algorithm is concentrated in `hf.py`, with focused footer tests and documentation. Changing the default behavior could affect callers and existing tests that expect convergence, so compatibility and API design—not code volume—are the main sources of risk.

## LS-08 — Upstream `SlideLayouts.get_by_name()` returns the first duplicate

### Severity and category

**Severity:** Low to medium
**Category:** Silent ambiguous selection in an inherited public API
**Mutation risk:** Indirect; callers may later apply the wrong returned layout

### Where it occurs

- `src/pptx/slide.py`: `SlideLayouts.get_by_name()`

### The lossy behavior

The method loops through layouts and returns the first layout whose name exactly matches the requested string. It does not detect duplicate names.

### What is bad about it

The method name implies lookup, but layout names are not guaranteed unique. A direct API user can therefore receive a valid-looking but unintended layout based on collection order.

This is distinct from LS-02: the import resolver has its own cross-master matching logic. Even after LS-02 is fixed, direct use of this inherited method remains ambiguous.

### Negative example

A master contains two layouts named `Comparison`, created through template evolution. A caller obtains `prs.slide_masters[0].slide_layouts.get_by_name("Comparison")` and adds a slide. The returned layout is the obsolete first layout rather than the intended newer layout.

### Impact

- Direct client code can select the wrong layout.
- Collection reorder changes behavior.
- The result offers no indication that another exact match existed.
- Paper workflows that call this API without additional uniqueness checks inherit the risk.

### Provenance

This behavior already existed at `paper-base` and is inherited from upstream `python-pptx`. Git history attributes the implementation to upstream work, not to the Paper bootstrap or current Paper PRs.

### Recommended fix

Avoid silently breaking a mature inherited API unless the project is prepared for a compatibility change. The smallest safe Paper addition is a strict companion API, for example `get_unique_by_name()` or a Paper-level layout resolver, that:

- returns the layout only when exactly one match exists;
- raises a typed not-found error for zero matches;
- raises a typed ambiguity error for multiple matches.

All Paper-authored workflows should use the strict resolver. The inherited `get_by_name()` can remain for compatibility, with documentation warning that it is first-match. A future major version could consider strengthening its semantics.

### Fix surface

**Localized if additive; potentially broad if the inherited method changes semantics.**

Adding a strict helper affects `slide.py`, focused tests, and Paper workflow call sites. Changing `get_by_name()` itself could break downstream `python-pptx`-compatible code and would require a larger compatibility review.

## Assessment of the current PR stack

The open PR stack does not introduce XLSX-style lossy selection or any of the runtime matchers above.

### PR #36 — OPC preservation

PR #36 changes package-level preservation and content-type handling. Its target normalization and relationship logic compare package semantics; they do not choose visible slide objects by fuzzy, positional, or geometric heuristics.

**Selection conclusion:** No identified contribution to lossy object selection.

### PR #37 — Table editing

PR #37 adds table editing APIs using explicit row and column indices and live cell proxies. Those are caller-specified coordinates in a table model, not inferred selection heuristics.

The APIs do not scan nearby cells, use character windows, or guess a cell from visual proximity.

**Selection conclusion:** No identified contribution to lossy selection. Exact index bounds and live-proxy semantics are the appropriate model here.

### PR #38 — Documentation and hygiene

PR #38 changes documentation, edit docstrings, and test harness material. It does not change runtime selection behavior.

Its relevant risk is descriptive: documentation about anchor safety should distinguish content freshness from stable object identity. Until LS-01 is fixed, docs should warn that duplicate text plus structural reordering can defeat the current stale check.

**Selection conclusion:** No runtime contribution; documentation should be corrected or qualified.

## Safe patterns already present

The package contains several selection designs worth reusing.

### Exact name lookup with ambiguity refusal

Shape-tree helpers such as `shape_by_name`, `picture_by_name`, `table_by_name`, and `chart_by_name` perform exact, group-aware lookup and raise `AmbiguousTargetError` when duplicate names prevent unique selection.

This is the right default pattern:

- exact candidate collection;
- unique match succeeds;
- zero matches report not found;
- multiple matches refuse and expose ambiguity.

### Exact text inspection

`inspect_text` traverses XML deterministically and documents blind regions rather than guessing at inaccessible text. It does not infer a target from nearby characters or pixels.

### Effective-value resolution

The effective-property resolver uses exact placeholder identity and refuses ambiguous resolution. Its “nearest” terminology refers to explicit OOXML inheritance precedence—not geometric proximity or fuzzy object matching.

### Conservative `refind()` behavior

When invoked, `refind()` searches for exact content fingerprints in the same part and refuses zero or multiple matches. Its limitation is that normal anchor resolution may falsely appear valid before refinding is attempted, as described in LS-01.

### Exact table coordinates

The table APIs introduced in PR #37 use explicit row/column coordinates and live cell proxies. They do not infer cells from spatial proximity.

## What the audit did not find

The Paper-added PPTX code did not reveal selectors based on:

- a fixed number of characters before or after a search hit;
- bounding-box overlap thresholds;
- nearest-object Euclidean distance;
- arbitrary pixel radii;
- fuzzy or Levenshtein name matching;
- “closest text” or “closest shape” geometry;
- selecting every object inside an approximate visual area.

Occurrences of words such as “nearest” in the inspected code refer to exact XML ancestry or inheritance precedence. They are not lossy object selection.

The absence of those heuristics is important: the recommended fixes should preserve that property. The answer to ambiguity should generally be stable identity, explicit caller choice, or refusal—not a new similarity score.

## Recommended remediation order

1. **Fix LS-01 first.** It can silently write to the wrong object despite the API's safety guard.
2. **Fix LS-02 and LS-03 next.** Both can silently apply wrong layout semantics during high-level structural operations.
3. **Fix LS-04 and LS-05 together where practical.** They share the principle that reports should use stable identity or conservative alignment rather than position.
4. **Fix LS-06.** It is a small change with a clear correctness benefit.
5. **Make an explicit product decision for LS-07.** The implementation is deterministic, but the default destructiveness deserves conscious approval.
6. **Add the strict API for LS-08.** Prefer an additive Paper-level safety API before changing inherited compatibility behavior.

## Cross-cutting design rules for fixes

The following rules address the root cause without overfitting to individual evaluation tasks:

1. **Stable structural identity before content.** Use shape IDs, group ID lineages, table coordinates, placeholder indexes, part names, and section IDs where the format provides them.
2. **Content as validation, not sole identity.** Text hashes are useful for stale detection but cannot distinguish identical content in different objects.
3. **Uniqueness at every fallback tier.** A fallback may proceed only when it yields exactly one candidate.
4. **Ambiguity is an explicit result.** Raise a typed error or report an unmatched/ambiguous record; do not silently choose by order.
5. **No approximate visual inference by default.** Avoid pixel distance, overlap, fuzzy text, or nearest-neighbor selection for mutating operations.
6. **Keep automatic and explicit paths separate.** Automatic inference should remain conservative; explicit mappings can provide the caller's missing intent.
7. **Use one shared exact alignment primitive.** Diff, rebind, and effective-format reporting should not develop inconsistent paragraph matching rules.
8. **Report destructive normalization.** If an operation deliberately deletes duplicates, make that an explicit mode and enumerate affected object IDs.
9. **Version persisted identities.** Changes to anchors or serialized match keys need schema versions and conservative legacy handling.
10. **Test reorder and duplicate cases.** Every selector should be tested with duplicate names/text, changed XML order, insertions, deletions, and groups—not only happy-path unique decks.

## Decision matrix

| Finding | If left unchanged | Benefit of fixing | Main compatibility concern |
|---|---|---|---|
| LS-01 | Rare but severe silent wrong text edits | Restores trustworthy inspect-then-edit safety | Stored v1 anchors and serialized outputs |
| LS-02 | Wrong theme/layout binding in ambiguous templates | Imports fail safely and become reproducible | More callers must provide `target_layout` |
| LS-03 | Content can jump to arbitrary compatible slots | Rebind/adopt-theme becomes intent-preserving | More callers must provide `placeholder_map` |
| LS-04 | Diff reports false shape movement/change | Validation reflects actual lineage edits | Golden/report keys may change |
| LS-05 | One paragraph insertion appears as many edits | Cleaner, more truthful text reports | Report schema and alignment goldens |
| LS-06 | Duplicate section names misroute imports | Small, clear correctness improvement | Callers with ambiguous decks must use ID |
| LS-07 | Footer operations can delete intentional duplicates | Prevents unexpected destructive cleanup | Existing convergence behavior may be relied on |
| LS-08 | Direct callers inherit first-match lookup | Provides a safe layout lookup path | Changing inherited API would affect compatibility |

## Bottom line

The package does not have a pervasive proximity-selection architecture and does not need a broad selection rewrite. The problems are concentrated in a small number of helpers that use order as a substitute for identity or uniqueness.

The seriousness varies sharply:

- `BlockAnchor` is a correctness and data-integrity problem because it can silently edit the wrong object.
- Layout and placeholder matchers are structural correctness problems because they silently bind to arbitrary compatible presentation structures.
- Diff matchers are observability problems because they can tell maintainers and evaluators the wrong story about a deck change.
- Section and inherited layout-name lookups are small ambiguity bugs.
- Footer canonicalization is primarily a product-policy decision because the loss is deliberate but potentially surprising.

The recommended direction is consistent across all findings: preserve stable identity where available, require uniqueness for automatic selection, and refuse ambiguity rather than replacing one weak heuristic with another.

---

This document is an audit and decision aid only. It does not modify runtime behavior or commit to a compatibility policy.
