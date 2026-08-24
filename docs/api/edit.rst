.. _edit_api:

Anchored text editing (``pptx.edit``)
=====================================

*paper-pptx addition.* Change text while preserving run formatting. Current paragraph anchors from
:func:`pptx.inspect.inspect_text` identify their owning shape or table cell structurally and then
validate a full content fingerprint. A changed fingerprint raises |StaleAnchorError| rather than
redirecting the edit. :func:`refind` is the explicit, exact recovery path.

Matching and boundaries
-----------------------

Both replacement functions use literal, case-sensitive, non-overlapping matches, processed from
left to right. ``find`` must be a non-empty string; ``replace`` is a string and may be empty. Tabs
are valid text. Line breaks (CR, LF, and vertical tab), other XML control characters, and text that
cannot be encoded for XML are refused before mutation.

A match may span adjacent text runs in one paragraph. It never crosses a paragraph boundary or an
intervening non-run element, including ``a:br`` line breaks and ``a:fld`` fields. This remains true
when the concatenated visible text appears to contain ``find`` across that boundary.

Formatting preservation
-----------------------

Fragments before and after a match keep the run properties of their source runs. Replacement text
uses the run properties of the run where the match begins. Runs untouched by a match remain
byte-identical. Any run left with no text after replacement is removed, so formatting that existed
only on that run is intentionally lost.

The caller therefore chooses the semantic span and which neighboring formatting remains. Suppose
``"Draft total"`` has ``"Draft "`` in regular text and ``"total"`` in bold. Replacing the whole
span with ``"Final total"`` gives the replacement the regular formatting at the start of the
match. Replacing only ``"Draft"`` with ``"Final"`` preserves the separate bold ``"total"`` run.

Results and refusal
-------------------

:func:`replace_text` traverses slide shapes, grouped shapes, and table cells, and optionally
existing notes slides. No match is a successful :class:`ReplaceResult` with zero replacements.
Unsupported markup-compatibility regions are blind to complete traversal, so their presence
refuses the deck-wide operation before any write.

:func:`replace_text_at` limits replacement to one anchored paragraph. It first resolves the named
slide or notes part, the slide-unique owning shape ID, table coordinates where applicable, and the
paragraph's container-local index. Only then does it validate the fingerprint. A missing or
ambiguous owner, changed table topology, invalid local paragraph, or duplicate exact fingerprint
inside the resolved container refuses before mutation; the diagnostic part-wide block ordinal is
never a current write target. A changed fingerprint raises |StaleAnchorError|. A missing literal
match (including one found only across a prohibited boundary) raises |TargetNotFoundError|.

An anchor addressing an unsupported blind region also refuses without mutation; unrelated blind
regions elsewhere do not block the anchored edit. Notes anchors returned by
``replace_text(..., include_notes=True)`` follow the same structural and fingerprint rules after
save and reopen. On success both replacement functions return a version-2
``paper-replace-result`` containing the count and fresh current anchors for the changed blocks.

:func:`refind` searches a current anchor only inside its structurally identified container for an
exact unique full-fingerprint match, refreshing local and diagnostic coordinates. A legacy
three-field anchor is searched exact-uniquely across its named part; its stored block ordinal is
ignored. No path uses fuzzy text, shape names, group order, geometry, or nearest-match behavior.

.. currentmodule:: pptx.edit

.. autofunction:: replace_text

.. autofunction:: replace_text_at

.. autofunction:: refind

.. autoclass:: ReplaceResult()
   :members:
   :undoc-members:
   :member-order: bysource
