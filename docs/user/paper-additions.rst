.. _paper_additions:

The paper-pptx additions
========================

**paper-pptx** is an agent-first, strict-superset hard fork of python-pptx. The distribution is
renamed; the import name stays ``pptx``. ``from pptx import Presentation`` and every other
existing call keep working unchanged. The added APIs cover four groups of operations:
**perceive**, **edit**, **compose**, and **verify**.

Do not install ``paper-pptx`` alongside ``python-pptx``: both distributions own the same
``pptx`` import package. Uninstall the upstream distribution before installing this fork.

This page summarizes the added APIs. Each capability's exact signatures, return types, and refusal
conditions live on the API pages linked throughout (and are collected under *paper-pptx
additions* in the :ref:`API Documentation <api>` table of contents).


Safety contract
---------------

The fork exists to prevent **silent corruption**: a deck that opens fine and is quietly wrong.
Every added operation either does exactly what it claims or refuses atomically. Mutating
operations follow **validate-fully-then-mutate**. When one cannot proceed safely, it raises a
typed refusal and leaves the document byte-for-byte unchanged in memory and on disk. The refusal
indicates that the operation was not applied.

The refusals form a small hierarchy rooted at |PaperRefusal| (see :ref:`errors_api`):
|PackageLimitError|, |TargetNotFoundError|, |AmbiguousTargetError|, |UnsupportedStructureError|,
|RelationshipPolicyError|, |BoundaryViolationError|, and |StaleAnchorError|. Programmer mistakes
(a bad type, an out-of-range index) stay plain ``ValueError`` / ``TypeError``. Callers can catch
``PaperRefusal`` separately::

    from pptx import Presentation
    from pptx.errors import PaperRefusal

    prs = Presentation("deck.pptx")
    try:
        prs.slides.clone(3)                 # slide contains an embedded OLE object
    except PaperRefusal as exc:
        ...                                 # document is untouched; handle or report exc

Normal package intake uses the same refusal boundary. It rejects ambiguous or unsafe ZIP
members before parsing XML, and refuses a package in which any member resolves to no content
type — no ``Default`` for its extension and no ``Override`` for its name in
``[Content_Types].xml`` — because PowerPoint refuses that package too. A part that nothing
references is accepted: PowerPoint opens such a package and drops the part on its next save,
which is also what ``save()`` does. A path-based ``save()`` writes beside the destination and
replaces it atomically only after serialization succeeds; stream saves retain normal stream
semantics.

Validation runs per mutating call by default. :meth:`~pptx.presentation.Presentation.batch`
trades that granularity for speed: inside the block the whole-deck check runs once, at exit,
rather than after every call. The check itself is unchanged, and it additionally covers the
mutation surface inherited from ``python-pptx``, which runs no transaction of its own — so a
sequence that previously saved an unreadable deck refuses instead. A refusal discards every
edit in the block, so scope a block to work you would be willing to redo, and save after it
closes (saving inside an open block raises |BoundaryViolationError|)::

    with prs.batch():
        for slide in prs.slides:
            slide.shapes[0].text_frame.text = "..."
    prs.save("deck.pptx")


Perceive a deck
---------------

Stock python-pptx returns ``None`` for any run property that is inherited rather than set
locally. On a branded template, that is nearly everything. :mod:`pptx.inspect` resolves those
values through the full inheritance walk (run → paragraph → shape list style → placeholder →
layout → master text styles → theme, with theme colors mapped through the master's ``clrMap``),
then reports *where each value came from*. Values that cannot be resolved are marked unresolved.

.. highlight:: python

::

    run = prs.slides[0].shapes.title.text_frame.paragraphs[0].runs[0]
    font = run.effective_font()             # -> EffectiveFont
    if font.size.resolved:
        print(font.size.value_pt)           # e.g. 36.0
        for step in font.size.provenance:   # the ordered chain of sources consulted
            if step.supplied:
                print("supplied by", step.level)   # e.g. "master txStyles titleStyle lvl1"

What cannot be resolved (a gradient fill, a missing theme) is reported as ``resolved=False`` with
its provenance intact. See :ref:`inspect_api` for
:func:`~pptx.inspect.effective_paragraph_format` and
:func:`~pptx.inspect.effective_shape_format` as well.

:func:`~pptx.inspect.effective_paragraph_format` resolves alignment, line spacing, and the
paragraph's **bullet** over that same inheritance walk, each reported with its own provenance. The
bullet arrives as an :class:`~pptx.inspect.EffectiveBullet` naming which of the four kinds actually
renders — ``"none"``, ``"character"``, ``"numbered"``, or ``"picture"`` — plus the glyph for a
character bullet, or the numbering scheme and starting number for a numbered one. Two answers
that read like absences are in fact resolved: a template that explicitly sets *no bullet* resolves
to ``"none"``, attributed to the rung that set it, and a chain with no bullet anywhere also resolves
to ``"none"``, through a closing ``rendering default`` provenance step. A paragraph inheriting the
master's ``•`` is therefore distinguishable from one that renders nothing.

The bullet's **typeface** and **size** arrive as two further values, ``bullet_font`` and
``bullet_size``. They are separate XSD choice groups that inherit on their own chains, so each
carries its own provenance: a paragraph can take its glyph from the master and its size from the
layout. The typeface matters more than it sounds — branded templates routinely use a symbol font,
where the glyph is a private-use codepoint that renders as a filled square in Wingdings and as an
empty box in anything else. Reporting both means a caller can reproduce a bullet from the payload
alone. Where the schema says to defer to the run's own text (``a:buFontTx``, ``a:buSzTx``), or where
no rung specifies anything, the value is |BULLET_FOLLOWS_TEXT| and ``resolved`` is still true. A
typeface naming a theme token such as ``+mj-lt`` resolves through the master's font scheme. The
enclosing ``.to_dict()`` payload is version 3 of ``paper-effective-paragraph-format``. ``"picture"`` is
reported on read only; there is no picture-bullet writer. Writing bullets is a separate job, done
with the |BulletFormat| authoring API described under **Edit one deck** below, which reads and
writes the paragraph's own markup.

Two functions emit deterministic, schema-versioned payloads (dataclasses with ``.to_dict()``)
built for diffing, golden-file tests, and automation:

* :func:`~pptx.inspect.inspect_text` — every text block on a slide, visibility-complete (it sees
  inside grouped shapes and table cells, which ordinary traversal skips), each block
  carrying a content-hash |BlockAnchor|.
* :func:`~pptx.inspect.inspect_deck` — a whole-deck structural manifest: per-slide shape
  inventory, geometry, placeholder roles, and layout/master inventory.


Edit one deck
-------------

**Anchored text replacement** (:ref:`edit_api`). :func:`~pptx.edit.replace_text` changes text
across the deck while preserving each run's formatting. :func:`~pptx.edit.replace_text_at`
targets one block by the |BlockAnchor| from :func:`~pptx.inspect.inspect_text`. The anchor carries
a hash of the block's text, so an edit aimed at content that has since changed raises
|StaleAnchorError| rather than landing in the wrong place. :func:`~pptx.edit.refind` is the
explicit recovery path.

**Slide operations** (:ref:`slides_api`). :meth:`~pptx.slide.Slides.clone`,
:meth:`~pptx.slide.Slides.delete`, :meth:`~pptx.slide.Slides.reorder`, and
:meth:`~pptx.slide.Slides.move` are in-memory, relationship-safe versions of operations that
previously required direct zip-package edits. Clone deep-copies charts *with their embedded
workbooks* and notes, so editing the clone's chart leaves the original byte-identical. It shares
media and refuses (|RelationshipPolicyError|) on relationship types it cannot safely honor.
Delete cannot leave orphaned parts and keeps section and custom-show lists consistent. The clone
policy is a |SlideClonePolicy|.

**Shape and table operations** (:ref:`shape_api`, :ref:`table_api`).
:meth:`~pptx.shapes.shapetree.SlideShapes.delete` /
:meth:`~pptx.shapes.shapetree.SlideShapes.move` /
:meth:`~pptx.shapes.shapetree.SlideShapes.add_copy`, plus group-aware by-name addressing
(:meth:`~pptx.shapes.shapetree.SlideShapes.shape_by_name`,
:meth:`~pptx.shapes.shapetree.SlideShapes.picture_by_name`,
:meth:`~pptx.shapes.shapetree.SlideShapes.table_by_name`,
:meth:`~pptx.shapes.shapetree.SlideShapes.chart_by_name`) refuse ambiguous names rather than
guess. On tables, :meth:`~pptx.table.Table.insert_row` /
:meth:`~pptx.table.Table.delete_row` / :meth:`~pptx.table.Table.insert_column` /
:meth:`~pptx.table.Table.delete_column` keep the grid definition consistent and guard merged
regions cell-wise. A merged header row does not block body-row operations.

Choose lookup by ownership. ``table_by_name()`` returns table content directly. If the edit also
needs width, height, position, or other geometry, retain the owning graphic frame with
``shape_by_name()``, check :attr:`~pptx.shapes.graphfrm.GraphicFrame.has_table`, and then use its
:attr:`~pptx.shapes.graphfrm.GraphicFrame.table`. Geometry belongs to the frame, not to |Table|::

    from pptx import Presentation
    from pptx.package import patch_save

    prs = Presentation("input.pptx")
    slide = prs.slides[0]
    frame = slide.shapes.shape_by_name("Plan Matrix")
    if not frame.has_table:
        raise ValueError("Plan Matrix is not a table")
    table = frame.table

    inserted_column_idx = 2
    table.insert_column(1, copy_format_from=1)
    table.cell(0, inserted_column_idx).text = "Current"
    delta = patch_save("input.pptx", prs, "output.pptx")

The returned |PackageDiff| is the residual package delta after unchanged original bytes have been
restored. Column insertion recalculates the owning frame's width; other shape geometry remains on
``frame``.

``copy_format_from`` names a pre-insertion column. It copies each source cell's direct formatting
into an empty, unmerged cell in the corresponding row; it does not materialize appearance inherited
from a table style or theme. An explicit ``width`` wins over the template width. Use
:meth:`~pptx.table._Cell.merge` to create a merge, :meth:`~pptx.table._Cell.extend_merge` to grow an
existing rectangular merge rightward, downward, or both, and :meth:`~pptx.table._Cell.split` to
remove one. Merge extension validates the entire requested rectangle before mutation, so malformed
topology, conflicts, and stale cells refuse without a partial edit.

**Text, notes, images, and charts** (:ref:`text_api`, :ref:`shape_api`, :ref:`chart-api`). Real
bullets and numbering via :attr:`~pptx.text.text._Paragraph.bullet` (a |BulletFormat|); autofit
you can read and freeze with :meth:`~pptx.text.text.TextFrame.normalize_autofit`; speaker-notes
:meth:`~pptx.slide.Slide.read_notes_text` / :meth:`~pptx.slide.Slide.replace_notes_text` that
leave absent notes parts absent; :meth:`~pptx.shapes.picture.Picture.replace_image` that swaps
pixels while keeping position, size, and crop byte-exact (optionally across formats); and
:meth:`~pptx.chart.chart.Chart.replace_data_safe`, which validates before it touches anything and
handles workbook-less charts, including those in the LibreOffice fixture corpus.

**Package comparison and narrow saves** (:ref:`package_api`).
:func:`~pptx.package.diff_package` reports what changed between two files, part by part, using
semantic XML comparison. Indentation is noise; a trailing space inside a text run is content. Valid
relationship registries compare complete bindings rather than child order, and valid content-type
registries compare the effective assignments for package members. Unknown or ambiguous registry
structures fall back conservatively; real relationship, content-type, member, and ordinary XML
ordering changes remain visible.
:func:`~pptx.package.patch_save` writes the edit and restores original bytes for every part that
did not semantically change. A one-line edit to a sixty-slide deck therefore diffs as one part,
not sixty. Table and shape APIs edit an in-memory |Presentation|; use ``save()`` for normal
serialization or ``patch_save(original_path, prs, out_path)`` when unchanged package members should
retain their original bytes.


Compose across decks
--------------------

Production decks are often assembled from many sources:
a pitch book's bank-overview pages may come from the master deck, its tombstones from the
credentials library, and its sector pages from the sector team. That relationship and
inheritance work is where decks get corrupted. The workflow is **import → renumber → deliver →
diff**.

**Real fields and footers** (:ref:`hf_api`). :meth:`~pptx.presentation.Presentation.apply_footers`
and :meth:`~pptx.slide.Slide.apply_footers` reproduce what PowerPoint's Insert → Header & Footer
dialog does, writing slide numbers and dates as genuine ``a:fld`` fields rather than static text.
The package authors the fields. PowerPoint or LibreOffice refreshes their values on open, so a
slide number written this way stays correct after a reorder.

**Layout rebind** (:ref:`rebind_api`). :meth:`~pptx.slide.Slide.rebind_layout` moves a slide to
another layout, matching placeholders by type and
index with an explicit map and orphan policy for the rest. Its |RebindReport| is required. The
resolver runs before and after, and every run whose *resolved* appearance changed is reported.

**Slide import and deck merge** (:ref:`compose_api`).
:meth:`~pptx.presentation.Presentation.import_slide` and
:meth:`~pptx.presentation.Presentation.append_deck` import a slide (or a whole deck) from another
presentation under one of three explicit reconciliation modes: ``"adopt_theme"`` (rebind to the
destination theme and report shifts), ``"keep_appearance"`` (transplant the source
layout/master/theme chain, hash-deduplicated so ten slides from one source do not create ten
masters), or ``"bake"`` (freeze effective values into explicit properties). The source
presentation remains unchanged. Charts travel with their workbooks, and unresolvable
relationships raise a typed refusal. Each import returns an |ImportReport|.

**Send-safe delivery.** Stripping speaker notes, review comments, and authorship before a deck
leaves the building needs no dedicated API: a part leaves the package by becoming unreachable, and
the serializer never writes an unreachable part. Drop the relationship and the part is gone::

    from pptx import Presentation
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    prs = Presentation("deck.pptx")
    for slide in prs.slides:
        for rId, rel in list(slide.part.rels.items()):
            if not rel.is_external and rel.reltype in (RT.NOTES_SLIDE, RT.COMMENTS):
                slide.part.drop_rel(rId)          # the notes/comments part stops being written
    prs.core_properties.author = ""
    prs.core_properties.last_modified_by = ""
    prs.save("clean.pptx")

Unused layouts go through :meth:`~pptx.slide.SlideLayouts.remove`, which refuses any layout a
slide still uses::

    for master in prs.slide_masters:
        for layout in list(master.slide_layouts):
            if not layout.used_by_slides:
                master.slide_layouts.remove(layout)

Verify what changed
-------------------

**Deck diff** (:ref:`diff_api`). :func:`~pptx.diff.diff_decks` reports slides added, removed, or
**moved** (matched by permanent slide id, so a
reorder is reported as a move rather than delete-plus-add), and within matched slides the shape,
chart-data, image, and notes deltas. At ``detail="full"``, it also reports per-run
effective-value shifts.
Callers can use the result to check a session's changes. Release job evaluations compare the
operation report with ``diff_decks(input, output)`` for every import/rebind/refresh job and
require them to agree.

A composition end to end::

    import shutil, tempfile, os
    from pptx import Presentation
    from pptx.diff import diff_decks

    tmp = tempfile.mkdtemp()
    house = shutil.copy("house_library.pptx", os.path.join(tmp, "house.pptx"))
    before = shutil.copy(house, os.path.join(tmp, "before.pptx"))

    prs = Presentation(house)
    source = Presentation("sector_team_deck.pptx")

    report = prs.import_slide(source, 0, mode="adopt_theme")   # -> ImportReport
    for shift in report.run_shifts:                            # appearance changes, reported
        print(shift.text, shift.before["name"]["value"], "->", shift.after["name"]["value"])

    prs.apply_footers(footer="Confidential", slide_number=True)   # real a:fld fields
    prs.core_properties.author = ""                               # send-safe, per above
    prs.save(os.path.join(tmp, "after.pptx"))

    delta = diff_decks(before, os.path.join(tmp, "after.pptx"), detail="text")
    print("slides added:", [s.slide_id for s in delta.slides_added])   # exactly the imported one
