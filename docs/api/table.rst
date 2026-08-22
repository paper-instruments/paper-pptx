.. _table_api:

Table-related objects
=====================


.. currentmodule:: pptx.table


|Table| objects
----------------

A |Table| object is added to a slide using the
:meth:`~.SlideShapes.add_table` method on |SlideShapes|.

Table content and shape geometry have different owners. Group-aware
:meth:`~pptx.shapes.shapetree.SlideShapes.table_by_name` returns a |Table| directly and is the
shortest path for cell, row, and column work. When the workflow also needs position, width, height,
rotation, or other shape properties, use
:meth:`~pptx.shapes.shapetree.SlideShapes.shape_by_name`, verify the returned graphic frame has a
table, and access its :attr:`~pptx.shapes.graphfrm.GraphicFrame.table`. Geometry stays on that
graphic frame; |Table| does not provide owner-navigation or geometry aliases.

.. autoclass:: Table()
   :members:
   :inherited-members:
   :exclude-members:
      notify_height_changed, notify_width_changed, part
   :undoc-members:


Direct-format column templates
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

:meth:`Table.insert_column` can copy direct cell formatting from an existing column. The source
index always names a column in the table *before* insertion. Each inserted cell receives the
complete direct cell-properties subtree from the source cell in the same row, including fills,
borders, margins, anchors, and producer extension content. Text and merge state are not copied;
the inserted cells are empty and unmerged.

This is a direct-format copy, not an effective-format calculation. Appearance supplied only by the
table style or theme is not materialized into the new cells. If the caller supplies a width, it
wins. Otherwise a selected source column supplies the new width; without a source, the existing
neighbor-width behavior applies.


|_Column| objects
-----------------

.. autoclass:: _Column()
   :members:
   :member-order: bysource
   :undoc-members:


|_Row| objects
--------------

.. autoclass:: _Row()
   :members:
   :member-order: bysource
   :undoc-members:


|_Cell| objects
---------------

A |_Cell| object represents a single table cell at a particular row/column
location in the table. |_Cell| objects are not constructed directly. A
reference to a |_Cell| object is obtained using the :meth:`Table.cell` method,
specifying the cell's row/column location. A cell object can also be obtained
using the :attr:`_Row.cells` collection.

.. autoclass:: _Cell
   :members:
   :member-order: bysource
   :undoc-members:


Creating, extending, and splitting merges
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

:meth:`_Cell.merge` creates a new rectangular merge from unmerged cells.
:meth:`_Cell.extend_merge` grows an existing merge-origin cell rightward, downward, or in both
directions in one rollback-safe operation. :meth:`_Cell.split` removes an existing merge.

Merge extension requires the receiver to be the live top-left origin of a canonical rectangular
merge and the target to be its requested new lower-right corner. It refuses malformed continuation
topology, stale cells, a target in another table, shrinking or upward/leftward movement, and any
other merge that intersects the requested rectangle. An unrelated merge elsewhere in the table
does not block the operation.

Only newly absorbed cell content moves to the origin, in row-major order after the origin's
existing content. The origin cell and its direct formatting remain in place, and table/frame
geometry does not change. Requesting the current lower-right corner is a no-op.
