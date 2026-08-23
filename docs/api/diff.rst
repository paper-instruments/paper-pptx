.. _diff_api:

Deck diff (``pptx.diff``)
=========================

*paper-pptx addition.* :func:`diff_decks` compares two decks and returns a typed report. It lists
slides added, removed, or **moved** (matched by permanent slide id, so a reorder is reported as a
move rather than delete-plus-add) and the shape, text, chart-data, image, and notes changes within
matched slides. At ``detail="full"`` it also reports per-run effective-value shifts and
per-paragraph **bullet shifts** via the resolver. Bullets live only in formatting, so a list losing
its bullets changes no text and no field marker: without that facet the within-slide report is empty
while the slide visibly loses every glyph. The report also includes ``package_changes``, an
authoritative semantic diff of every package member. That package-level list prevents
metadata, relationship, ordering, field, crop, or media changes from disappearing when no
specialized slide facet applies.

Contract tests compare operation reports with ``diff_decks(input, output)`` for representative
import, rebind, and refresh workflows.

Matching uses the permanent slide id, which serves lineage-derived decks (v4 saved from v3).
Within a matched slide, top-level shapes match by their slide-wide ``shape_id`` and compatible
structural kind. A shape's name is display metadata only: duplicate, empty, or renamed shapes keep
their identity, and changing z-order alone produces no specialized shape addition, removal,
geometry, image, or chart entry. The real XML-order change remains available in
``package_changes``.

The top-level boundary is deliberate. Moving a leaf into a group reports a removal; moving one out
reports an addition. The diff does not recursively rescue that association. Independently built
decks can reuse numeric ids and are outside the lineage contract. Deleting a shape and later
reusing its id for a new shape of the same structural kind is also indistinguishable without a
persistent identifier that general PPTX files do not provide.

Text inspection is visibility-complete even though the public shape facets are intentionally
top-level. Ordinary text and recursively grouped leaf text use the inspection-visible leaf's exact
slide-wide shape ID and compatible structural kind as their text-container boundary. Table text
uses the table graphic-frame ID. Group display names, group order, shape geometry, and text
similarity are never identity.

``paper-deck-diff`` schema version 5 uses one structured shape reference everywhere. Entries in
``shapes_added``, ``shapes_removed``, and ``images_replaced`` are
``{"shape_id": ..., "name": ...}``; the ``shape`` value in ``geometry_changes`` and the ``chart``
value in ``chart_data_changes`` use that same structure. Additions use the after-side name,
removals use the before-side name, and matched facets use the after/current name. Consumers
migrating from version 3 should compare ``entry["shape_id"]`` (or
``entry["shape"]["shape_id"]`` / ``entry["chart"]["shape_id"]``) rather than the former string
label. A rename alone remains observable through ``package_changes`` and does not manufacture a
specialized shape change.

Text comparison is a deterministic snapshot comparison, not recovered edit history. Within each
stable container, paragraphs are ordered values made from raw literal text and positioned field
markers. Run boundaries and formatting do not participate. The matcher consumes the longest exact
prefix first and then the longest non-overlapping exact suffix. It emits at most one event for the
remaining middle:

* ``insertion`` when only the after range contains paragraphs;
* ``deletion`` when only the before range contains paragraphs;
* ``replacement`` when each range contains exactly one paragraph; or
* ``changed_region`` for every larger two-sided change, including paragraph reorders.

The comparison is exact: whitespace and Unicode representation are content, so canonically
equivalent but differently encoded strings remain different. There is no paragraph ``move`` or
``ambiguity`` event and no claim that equal text denotes one persistent paragraph. For repeated
values, consuming the prefix before the suffix gives one canonical hunk location. That location may
differ from where a user historically inserted or removed a duplicate paragraph.

An unmatched text-bearing container contributes one whole-container insertion or deletion event
alongside any shape addition or removal. It is not paired with another unmatched container.

All four kinds use the same version-5 payload::

    {
        "kind": "replacement",
        "shape": {"shape_id": 12, "name": "Summary"},
        "before_range": {"start": 0, "end": 1},
        "after_range": {"start": 0, "end": 1},
        "before": [
            {
                "location": {"container": "shape", "paragraph_index": 0},
                "text": "Old summary",
                "fields": [],
            }
        ],
        "after": [
            {
                "location": {"container": "shape", "paragraph_index": 0},
                "text": "New summary",
                "fields": [{"offset": 11, "type": "slidenum"}],
            }
        ],
    }

``before_range`` and ``after_range`` are zero-based, half-open paragraph indexes in that side's
container sequence. ``before`` and ``after`` are always arrays; an absent side is ``[]``, never
``null``. Every block has raw ``text``, an always-present ``fields`` array of positioned
``{offset, type}`` records, and its exact container-local ``location``. A table-cell location also
has ``row`` and ``column``; its paragraph index is local to that cell. Table container ranges index
the frame's complete row-major paragraph sequence, while block locations retain their snapshot
coordinates.

Version-4 consumers must replace scalar ``before``/``after`` values and
``shape_id``/``shape_name``/``block_ordinal`` keys with the regular version-5 arrays, ranges,
structured ``shape``, and per-block locations. Version 5 has no scalar compatibility form.

At ``detail="structure"``, ``table_structure_changes`` reports only stable table-frame identity
and exact before/after row and column counts. It does not infer which row or column was historically
inserted or deleted. At ``detail="text"``, the row-major hunk provides exact cell evidence but its
coordinates identify snapshot positions, not persistent cells; column growth can therefore produce
one broad ``changed_region``. At ``detail="full"``, effective-font and bullet comparison is limited
to an unchanged prefix/suffix paragraph whose complete text-and-field value occurs exactly once in
each container. Runs then match only by a non-empty exact text identity or exact field type that is
itself unique within each paragraph. Repeated paragraphs, changed-region paragraphs, table
effective formatting, and table bullets deliberately produce no specialized shift.

Speaker notes retain the separate flat ``notes_change`` comparison. ``package_changes`` remains
the authoritative semantic package-level fallback when a specialized facet is intentionally
coarse or unsupported, including formatting inside changed text and table-style formatting.

.. currentmodule:: pptx.diff

.. autofunction:: diff_decks

.. autoclass:: DeckDiff()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: SlideChange()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: ShapeRef()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: BulletShift()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: EffectiveShift()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: SlideRef()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: MovedSlide()
   :members:
   :undoc-members:
   :member-order: bysource
