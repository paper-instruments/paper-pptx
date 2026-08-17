# The oxml layer

`xmlchemy` is an object-XML mapping layer for `lxml`, loosely in the spirit of SQLAlchemy —
hence the name. It exists so that the very broad schema of Open XML elements can be
expressed declaratively instead of by hand-assembling elements.

**Never hand-assemble lxml elements in proxy or API code.** New XML vocabulary is declared
here, in the oxml layer, and surfaced through a thin proxy.

Everything below is checked against `xmlchemy.py` itself. It supersedes the Sphinx-era
`xmlchemy` page, which covered three of the eight descriptors.

## Adding support for a new element type

1. Add a custom element mapping in `pptx/oxml/__init__.py`.
2. Add a custom element class in the appropriate `pptx.oxml` subpackage module, subclassing
   `BaseOxmlElement`.
3. Declare its children with the element descriptors below.
4. Declare its attributes with the attribute descriptors below.
5. Add any new simple types to `pptx/oxml/simpletypes.py`.

Commit in that shape: analysis → oxml classes → proxy → tests.

```python
from pptx.oxml.xmlchemy import BaseOxmlElement, ZeroOrOne, OptionalAttribute

class CT_Foobar(BaseOxmlElement):
    """Custom element class for the CT_Foobar complex type in pml.xsd."""

    hlink = ZeroOrOne("a:hlink", successors=("a:rtl", "a:extLst"))
    sz = OptionalAttribute("sz", ST_SimpleType, default=ST_SimpleType.OPTION)
    rId = RequiredAttribute("r:id", XsdString)
```

## `successors=` is the schema's child order

Open XML requires children in a specific sequence, and `successors=` is this library's
encoding of that sequence: the list of tags that must come *after* the declared element. The
generated inserter places a new child before the first successor it finds.

Get this wrong and you produce a file that opens in this library and is rejected by
PowerPoint. **Do not port ordered-insertion tables from reference helpers into `src/`** —
that mechanism exists only because those helpers lived outside the library. Their knowledge
of which orderings matter is still useful review input.

The authoritative source for a given element's order is the xsd excerpt in its
`docs/dev/analysis/` page — see below.

## Element descriptors

For a declaration `child = Descriptor("ns:localTagName", successors=(...))`:

| Descriptor | Generated API |
|---|---|
| `OneAndOnlyOne` | read-only property. Raises `InvalidXmlError` on access when the required child is absent. Takes **no** `successors` — nothing is ever inserted. |
| `ZeroOrOne` | read-only property returning the child or `None`; `get_or_add_child()`; plus `_add_child()`, `_new_child()`, `_insert_child()`, `_remove_child()` for overriding. |
| `ZeroOrMore` | `child_lst` list property — **the declared name itself is deleted from the class**, so use `child_lst`; plus `_new_child()`, `_insert_child()`, `_add_child()`. |
| `OneOrMore` | as `ZeroOrMore`, plus a public `add_child()`. |
| `Choice` | used inside `ZeroOrOneChoice`. Gives `get_or_change_to_child()` and `_remove_child()`. |
| `ZeroOrOneChoice` | a mutually exclusive group of `Choice` members. The group property returns whichever member is present, or `None`; `_remove_group()` clears it. |

The private hooks exist to be overridden when an element needs customised creation or
insertion. Override the hook; do not bypass the descriptor.

```python
>>> foobar.hlink                      # ZeroOrOne, absent
None
>>> hlink = foobar.get_or_add_hlink()
>>> foobar.hlink is hlink
True

>>> foobar.eg_fillProperties          # ZeroOrOneChoice, absent
None
>>> solidFill = foobar.get_or_change_to_solidFill()
>>> foobar.eg_fillProperties is solidFill
True
>>> foobar.remove_eg_fillProperties()
```

## Attribute descriptors

Both take a tag name and a simple-type class, and generate a read/write property. Reading
type-converts through the simple type's `from_xml()`; assignment is validated by its
`validate()` and encoded by its `to_xml()`, so values are used in their natural Python type.
Invalid assignments raise `TypeError` or `ValueError`.

- `RequiredAttribute` — the attribute must be present. Reading it when absent raises
  `InvalidXmlError`.
- `OptionalAttribute` — accepts `default=`, returned when the attribute is absent.

## `r:id` is an indirect reference

An element that points at another part does not name it. It carries a part-local
relationship id — `r:embed` on `a:blip`, `r:id` elsewhere — and that id is resolved through
the source part's own `.rels` item:

```xml
<a:blip r:embed="rId2"/>                          <!-- in the slide -->
<Relationship Id="rId2" Type=".../image" Target=".../image1.png"/>   <!-- in its .rels -->
```

The indirection is deliberate. It lets a slide reference the image it displays without
knowing where that image will land in the saved package, which is what keeps packaging a
separate concern from content. So when you declare an attribute that refers to another part,
declare it as the relationship id — never a path.

Two corruption classes follow from getting this wrong, and `tests/paper/relint.py` scans
every output for both: a part referencing an `r:id` its `.rels` does not define, and a
`.rels` entry targeting a package member that does not exist. A relationship with a
reference count of zero is an *implicit* relationship and is legitimate; see
`XmlPart.drop_rel` in `pptx/opc/package.py`.

## Whitespace is content

Text nodes with preserved-space semantics — `a:t` and friends — must never be normalised by
any comparison, canonicalisation, or rewrite path. A canonicaliser that trims a meaningful
trailing space will make `patch_save` "restore" original bytes over a real edit, which is
corruption inside the safety tooling itself.

## Read the analysis notes first

**Before adding XML vocabulary, read the design note for that feature area in
`docs/dev/analysis/`.**

Those 79 files are python-pptx's own per-feature notes, inherited with the fork and never
modified. Each pairs prose with a real XML specimen and the matching xsd excerpt. They
describe the OOXML format and how this layer models it — the layer the fork does not
change — which is why they survived the removal of the rest of the Sphinx tree. They are not
published anywhere and are not a documentation target; they are working reference.

`txt-text.rst` is the pattern worth seeing first. Its prose explains that a paragraph is a
sequence of `a:r`, `a:br`, and `a:fld`; its specimen shows an `a:fld` carrying its own `a:t`
that PowerPoint regenerates on render — which is why text replacement must not cross a field
boundary, and why editing inside one silently loses the edit. Its xsd excerpt gives
`CT_TextBody`'s required child sequence, which is what a `successors=` list has to match.
