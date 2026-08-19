*paper-pptx* is Paper Instruments' agent-first hard fork of *python-pptx* ``v1.0.2`` for safely
**inspecting, editing, and composing** existing PowerPoint (.pptx) files. Its public Python API
is a superset of upstream's; package intake and save publication are deliberately stricter.
The distribution is renamed; the import name stays ``pptx``. ``from pptx import Presentation``
and every other existing call keep working unchanged.

Stock python-pptx provides the package and XML layers for creating decks. paper-pptx adds APIs for
existing decks: resolve effective values through the placeholder → layout → master → theme chain,
with provenance; edit text and structure; compose slides across files under an explicit theme
policy; and compare the result with a deck-to-deck diff. It runs on any Python-capable platform,
including macOS and Linux, and does not require PowerPoint to be installed or licensed.

The fork exists to prevent **silent corruption**: a deck that opens fine and is quietly wrong.
Every added operation either does exactly what it claims or refuses atomically, leaving the
document byte-for-byte unchanged. The contract harness checks this through save → reopen checks,
exact changed-part budgets, and an independent LibreOffice load smoke in release verification.

Start with :ref:`paper_additions` for an overview. Each added module has a page under
`API Documentation`_. The remaining documentation is inherited from python-pptx and describes
the shared, unchanged foundation.
