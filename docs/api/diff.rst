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

``paper-deck-diff`` schema version 5 uses one structured shape reference everywhere. Entries in
``shapes_added``, ``shapes_removed``, and ``images_replaced`` are
``{"shape_id": ..., "name": ...}``; the ``shape`` value in ``geometry_changes`` and the ``chart``
value in ``chart_data_changes`` use that same structure. Additions use the after-side name,
removals use the before-side name, and matched facets use the after/current name. Consumers
migrating from version 3 should compare ``entry["shape_id"]`` (or
``entry["shape"]["shape_id"]`` / ``entry["chart"]["shape_id"]``) rather than the former string
label. A rename alone remains observable through ``package_changes`` and does not manufacture a
specialized shape change.

Text comparison is likewise lineage-scoped. Paragraphs are partitioned by the slide-wide ID of
their ordinary/grouped leaf shape or table frame, then aligned by the exact inspection-v3
fingerprint (NFC-normalized literal segments plus field type and position). The matcher does not
use paragraph ordinal, display name, geometry, edit distance, or fuzzy text. Insertions and
deletions therefore do not shift unchanged neighbors into fictional replacements. Exact unique
reorders can be reported as ``kind="move"``; a uniquely bounded one-for-one change is a
``replacement``. Repeated regions that cannot be associated from exact context use one
``ambiguity`` event rather than an arbitrary pairing.

Every text event carries a structured ``shape`` reference and separate ``before_location`` and
``after_location``. A location contains the container kind and container-local paragraph index;
table-cell locations also contain row and column. Insertions and deletions use ``null`` on the
absent side. Ambiguity events keep the singular locations null and list the exact candidates.
Version-4 consumers must migrate from ``shape_id``/``shape_name``/``block_ordinal`` to these
version-5 fields.

At ``detail="structure"``, ``table_structure_changes`` reports stable table-frame identity and
before/after row and column counts. A provably located row or column insertion/deletion includes
its index; duplicate or blank cells retain candidate indexes and an ambiguity marker. At
``detail="text"`` the same frame-scoped alignment retains separate cell coordinates. At
``detail="full"`` effective-font and bullet shifts reuse those exact paragraph pairs and have
before/after locations. Table effective formatting and table bullets remain unsupported and are
not inferred. Speaker notes retain the separate flat ``notes_change`` comparison.

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
