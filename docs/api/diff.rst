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
authoritative semantic diff of every serialized package member. That package-level list prevents
metadata, relationship, ordering, field, crop, or media changes from disappearing when no
specialized slide facet applies.

Contract tests compare operation reports with ``diff_decks(input, output)`` for representative
import, rebind, and refresh workflows.

Matching uses the permanent slide id, which serves lineage-derived decks (v4 saved from v3).
Independently built decks can reuse the same numeric ids and are outside that matching contract.

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
