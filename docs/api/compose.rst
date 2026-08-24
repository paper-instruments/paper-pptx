.. _compose_api:

Slide import and deck merge (``pptx.compose``)
==============================================

*paper-pptx addition.* Import a slide or a whole deck from another presentation. Much of a
slide's appearance lives outside the slide, in its layout, master, and theme. Import is therefore
an *inheritance-reconciliation* problem with three explicit modes:

- **adopt_theme**: rebind the incoming slide to a unique exact-name destination layout, or when
  none exists, a unique exact non-custom-type layout, so it takes the destination theme;
  appearance shifts are included in the report.
- **keep_appearance**: transplant the source layout / master / theme chain, deduplicated by
  content hash so importing ten slides from one source does not create ten masters.
- **bake**: snapshot the slide's effective values into explicit properties, then attach through
  the same unique name/type tiers or a unique blank-layout fallback: visually stable without
  importing masters. Bake never falls back to the first destination layout.

Automatic tiers are evaluated in that order: an empty tier continues, while the first non-empty
tier must contain exactly one candidate. Multiple candidates at a stronger tier raise
|AmbiguousTargetError| before any destination write; the refusal lists each layout's name, type,
part, and owning master. Pass an enrolled destination ``target_layout`` to
:meth:`.Presentation.import_slide` to make the choice explicit. A whole-deck append preflights
every source slide under the same rules, so ambiguity on a later slide appends nothing.

Adopt-theme placeholder reconciliation uses the same exact-first, unique-only matcher as layout
rebind. Same-type and compatible-family fallbacks bind only when their current unclaimed candidate
set has one member; ambiguity raises |AmbiguousTargetError| without trying a weaker tier. Supply a
partial ``placeholder_map={source_idx: target_idx | None}`` to
:meth:`.Presentation.import_slide` to settle selected sources explicitly while the rest reconcile
automatically. ``None`` deliberately orphans that source placeholder, which adopt-theme bakes from
its source-resolved appearance. The argument is rejected for keep-appearance and bake. Whole-deck
append deliberately remains automatic-only and refuses atomically if any staged slide needs a map.

``paper-import-report`` version 2 always serializes ``placeholder_map_used``. Adopt-theme reports
the complete resolved source-to-target mapping, including ``null`` orphan targets, in source-index
order. Keep-appearance and bake serialize an empty list because they do not reconcile
placeholders.

The source presentation remains unchanged. Charts travel with their embedded workbooks, media is
always copied across packages, and relationships that cannot be resolved refuse
(|RelationshipPolicyError|).
Source and destination slide dimensions must match exactly. Import refuses rather than silently
rescaling or clipping content.

The entry points are methods on |Presentation|; see :meth:`.Presentation.import_slide` and
:meth:`.Presentation.append_deck`. This page documents the report they return.

.. currentmodule:: pptx.compose

.. autoclass:: ImportReport()
   :members:
   :undoc-members:
   :member-order: bysource
