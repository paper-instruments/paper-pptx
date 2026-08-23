.. _package_api:

Package kernel (``pptx.package``)
=================================

*paper-pptx addition.* Compare ``.pptx`` packages semantically and save edits narrowly.
Comparison reports changes part by part. A narrow save lets only changed parts differ
from the original.

Comparison treats structural whitespace between elements as noise but preserves meaningful text
whitespace, including trailing spaces inside text runs.

Valid OPC relationship parts compare by their complete relationship bindings rather than child
order. Omitted ``TargetMode`` and explicit ``TargetMode="Internal"`` are equivalent, as are accepted
absolute and source-relative spellings that resolve to the same internal member; this tolerance is
not a recommendation to author absolute internal targets. IDs, relationship types, whether a
binding is external, external target spellings, and unknown attributes remain significant. Valid
content-type manifests compare the effective content type assigned to every package member.
Unmatched default declarations remain significant. Overrides for absent members, missing or
ambiguous assignments, genuine assignment changes, and package member additions or removals remain
visible; unsupported manifest structures fall back to the existing declaration-order-insensitive
XML comparison.

All other XML remains order-sensitive. Unsupported relationship structures fall back to that
ordinary XML comparison; unsupported content-type manifests retain the existing
declaration-order-insensitive comparison. When comparison requires parsing, malformed XML and
prohibited DTD/entity constructs raise ``ValueError``.

:func:`diff_package`, :attr:`pptx.diff.DeckDiff.package_changes`, and
:func:`patch_save` use this same package-member verdict. The first returns the schema-version-1
|PackageDiff| directly, the deck diff exposes its deltas with the higher-level slide facets, and a
narrow save uses the verdict to restore original bytes before returning its residual
|PackageDiff|.

.. currentmodule:: pptx.package

.. autofunction:: xml_equivalent

.. autofunction:: diff_package

.. autofunction:: patch_save

.. autoclass:: PackageDiff()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: PartDelta()
   :members:
   :undoc-members:
   :member-order: bysource
