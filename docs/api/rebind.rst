.. _rebind_api:

Layout rebind (``pptx.rebind``)
===============================

*paper-pptx addition.* Move a slide to another layout in the same presentation, matching
placeholders by type and index with an explicit map and an orphan policy for the rest.

Automatic reconciliation first settles every exact type-and-index match across the slide. It
then considers same-type matches, followed by the compatible title and content families. A
fallback tier binds only when exactly one unclaimed target placeholder exists. Multiple candidates
raise |AmbiguousTargetError| before mutation and list each target type/index; pass a partial
``placeholder_map={source_idx: target_idx | None}`` to choose explicitly. Unlisted sources still
use automatic matching, and ``None`` deliberately orphans a source so ``orphan_policy`` controls
whether it refuses or bakes to a free shape.

The report is required. The effective-value resolver runs before and after, and every text run
whose *resolved* appearance changed appears in it. ``paper-rebind-report`` remains version 1 and
records the complete resolved mapping in ``placeholder_map_used``.

The entry point is a method on |Slide|; see :meth:`.Slide.rebind_layout`. This page documents the
report it returns.

.. currentmodule:: pptx.rebind

.. autoclass:: RebindReport()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: RunShift()
   :members:
   :undoc-members:
   :member-order: bysource
