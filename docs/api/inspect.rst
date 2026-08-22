.. _inspect_api:

Inspection (``pptx.inspect``)
=============================

*paper-pptx addition.* Resolve the values a deck actually renders, with provenance and without
mutation. These are the sizes, fonts, and colors that stock python-pptx returns as ``None``
because they are inherited through the placeholder → layout → master → theme chain. Every
resolved value explains where it came from. A value that cannot be resolved is reported as
unresolved.

.. currentmodule:: pptx.inspect


Functions
---------

.. autofunction:: effective_font

.. autofunction:: effective_paragraph_format

.. autofunction:: effective_shape_format

.. autofunction:: inspect_text

.. autofunction:: inspect_deck


Resolved values and provenance
-------------------------------

.. autodata:: BULLET_FOLLOWS_TEXT

.. autoclass:: EffectiveValue()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: ProvenanceStep()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: EffectiveFont()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: EffectiveBullet()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: EffectiveParagraphFormat()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: EffectiveShapeFormat()
   :members:
   :undoc-members:
   :member-order: bysource


Text inspection payloads
------------------------

``paper-text-inspection`` version 3 gives every block a current version-2 |BlockAnchor|. Its
``locator`` identifies a text shape by slide-unique shape ID and paragraph index within that
shape. Grouped leaves use the leaf shape ID directly; group names and order are display context,
not identity. Table cells add the graphic-frame shape ID, zero-based row and column, and paragraph
index within that cell. The part name supplies the notes-slide identity for notes anchors emitted
by the editing APIs.

The retained ``block_index`` is a diagnostic part-wide traversal ordinal only. Current writes
resolve the locator first and never use that ordinal as identity. ``content_hash`` on a current
anchor is a full SHA-256 fingerprint over NFC-normalized literal content plus ordered, positioned
field-type markers. Meaningful whitespace and hard line breaks remain content; cached field
display values do not. The public :func:`content_hash` helper is unchanged: it returns the legacy
eight-character NFC SHA-256 prefix and never trims whitespace.

Existing three-argument ``BlockAnchor(part, block_index, content_hash)`` construction creates a
legacy anchor. It carries no structural locator; editing can accept it only when that short hash
matches exactly one block in the named part.

.. autofunction:: content_hash

.. autoclass:: BlockAnchor()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: InspectedRun()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: TextBlock()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: TextInspection()
   :members:
   :undoc-members:
   :member-order: bysource


Deck manifest payloads
----------------------

.. autoclass:: DeckManifest()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: SlideManifest()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: ShapeManifest()
   :members:
   :undoc-members:
   :member-order: bysource
