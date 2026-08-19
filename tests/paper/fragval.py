"""Scoped schema validation of emitted XML fragments (the second oracle).

Validates exactly the fragments paper code writes — never whole documents, which drown in
upstream noise. Fragments are checked against the ECMA-376 transitional schemas shipped in
`tests/paper/xsd/` (the ISO/IEC 29500-4 edition whose namespaces match real-world files), via a tiny
generated wrapper schema that promotes the fragment's complex type to a global root element.
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict

from lxml import etree

_SPEC_XSD_DIR = Path(__file__).resolve().parent / "xsd"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

_WRAPPER_TEMPLATE = """\
<xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema"
    xmlns:a="{ns}" targetNamespace="{ns}" elementFormDefault="qualified">
  <xsd:include schemaLocation="{dml_main}"/>
  <xsd:element name="{root_local}" type="a:{complex_type}"/>
</xsd:schema>
"""

_schema_cache: "Dict[tuple, etree.XMLSchema]" = {}


def _schema_for(root_local: str, complex_type: str) -> etree.XMLSchema:
    """Return a compiled wrapper schema validating `<a:{root_local}>` as `{complex_type}`."""
    key = (root_local, complex_type)
    if key not in _schema_cache:
        wrapper = _WRAPPER_TEMPLATE.format(
            ns=_A_NS,
            dml_main=(_SPEC_XSD_DIR / "dml-main.xsd").as_uri(),
            root_local=root_local,
            complex_type=complex_type,
        )
        _schema_cache[key] = etree.XMLSchema(etree.fromstring(wrapper.encode()))
    return _schema_cache[key]


def _validation_errors(element, root_local: str, complex_type: str) -> "list[str]":
    schema = _schema_for(root_local, complex_type)
    # -- re-parse a clean serialization so validation can't be affected by proxy classes
    fragment = etree.fromstring(etree.tostring(element))
    if schema.validate(fragment):
        return []
    return [str(entry) for entry in schema.error_log]


def assert_pPr_fragment_valid(paragraph) -> None:
    """Assert the `a:pPr` fragment of a python-pptx paragraph validates against ECMA-376.

    No-op (vacuously true) when the paragraph has no `a:pPr`.
    """
    pPr = paragraph._p.pPr
    if pPr is None:
        return
    errors = _validation_errors(pPr, "pPr", "CT_TextParagraphProperties")
    assert not errors, "emitted a:pPr fragment fails schema validation:\n%s" % "\n".join(errors)


def assert_bodyPr_fragment_valid(text_frame) -> None:
    """Assert the `a:bodyPr` fragment of a text frame validates against ECMA-376."""
    errors = _validation_errors(text_frame._bodyPr, "bodyPr", "CT_TextBodyProperties")
    assert not errors, (
        "emitted a:bodyPr fragment fails schema validation:\n%s" % "\n".join(errors)
    )


def assert_tbl_fragment_valid(table) -> None:
    """Assert a table's complete `a:tbl` element validates against ECMA-376 CT_Table.

    The table oracle: row/column surgery rewrites `a:tr`/`a:tc`/`a:gridCol`
    structure, so the whole emitted table element is validated, not just one fragment.
    `a:tbl` is already a global element in dml-main.xsd, so no wrapper schema is needed.
    """
    key = ("__global__", "tbl")
    if key not in _schema_cache:
        _schema_cache[key] = etree.XMLSchema(etree.parse(str(_SPEC_XSD_DIR / "dml-main.xsd")))
    schema = _schema_cache[key]
    fragment = etree.fromstring(etree.tostring(table._tbl))
    assert schema.validate(fragment), (
        "emitted a:tbl fragment fails schema validation:\n%s"
        % "\n".join(str(entry) for entry in schema.error_log)
    )
