.. _edit_api:

Anchored text editing (``pptx.edit``)
=====================================

*paper-pptx addition.* Change text while preserving run formatting. Paragraph blocks are
addressed by the content-hash |BlockAnchor| from :func:`pptx.inspect.inspect_text`. Because the
anchor carries a hash of the block's text, an edit aimed at content that has since changed raises
|StaleAnchorError| rather than being silently misapplied. :func:`refind` is the explicit recovery
path.

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

:func:`replace_text_at` limits replacement to one anchored paragraph. A changed content hash raises
|StaleAnchorError|; a missing match (including one found only across a prohibited boundary) raises
|TargetNotFoundError|. An anchor addressing an unsupported blind region also refuses without
mutation; unrelated blind regions elsewhere do not block the anchored edit. On success both
functions return a :class:`ReplaceResult` containing the count and post-edit anchors for the blocks
that changed. Use :func:`refind` explicitly to recover the unique current location of stale
anchored content; it refuses when no block or more than one block has the requested content hash.

.. currentmodule:: pptx.edit

.. autofunction:: replace_text

.. autofunction:: replace_text_at

.. autofunction:: refind

.. autoclass:: ReplaceResult()
   :members:
   :undoc-members:
   :member-order: bysource
