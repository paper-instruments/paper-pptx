.. _text_api:

Text-related objects
====================


.. currentmodule:: pptx.text.text


|TextFrame| objects
--------------------

.. autoclass:: TextFrame()
   :members:
   :member-order: bysource
   :undoc-members:


|Font| objects
--------------

The |Font| object is encountered as a property of |_Run|, |_Paragraph|, and in
future other presentation text objects.

.. autoclass:: Font()
   :members:
   :member-order: bysource
   :undoc-members:


|_Paragraph| objects
--------------------

.. autoclass:: _Paragraph()
   :members:
   :member-order: bysource
   :undoc-members:


|_Run| objects
--------------

.. autoclass:: _Run()
   :members:
   :member-order: bysource
   :undoc-members:


|BulletFormat| objects
----------------------

*paper-pptx addition.* Read and set real bullet/numbering on a paragraph, accessed through the
:attr:`._Paragraph.bullet` property. Retires the fake-glyph anti-pattern (a literal "•" typed
into the text): these write genuine ``a:buChar`` / ``a:buAutoNum`` / ``a:buNone`` markup with
hanging-indent control. This object reads and writes the paragraph's own markup; to see which
bullet actually renders once the layout and master chain is taken into account, use the
effective-style inspection API in :ref:`inspect_api` —
:func:`~pptx.inspect.effective_paragraph_format` reports it as an |EffectiveBullet|.

.. autoclass:: pptx.text.bullet.BulletFormat()
   :members:
   :member-order: bysource
