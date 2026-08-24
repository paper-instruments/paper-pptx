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

``paper-deck-diff`` schema version 4 uses one structured shape reference everywhere. Entries in
``shapes_added``, ``shapes_removed``, and ``images_replaced`` are
``{"shape_id": ..., "name": ...}``; the ``shape`` value in ``geometry_changes`` and the ``chart``
value in ``chart_data_changes`` use that same structure. Additions use the after-side name,
removals use the before-side name, and matched facets use the after/current name. Consumers
migrating from version 3 should compare ``entry["shape_id"]`` (or
``entry["shape"]["shape_id"]`` / ``entry["chart"]["shape_id"]``) rather than the former string
label. A rename alone remains observable through ``package_changes`` and does not manufacture a
specialized shape change.

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

.. autoclass:: SlideRef()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: MovedSlide()
   :members:
   :undoc-members:
   :member-order: bysource
