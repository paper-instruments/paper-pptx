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
