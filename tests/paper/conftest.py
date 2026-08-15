"""Shared pytest configuration for the paper test suite (tests/paper only)."""

from __future__ import annotations

import os

import pytest

from pptx.opc.package import PartFactory

from .clock import FrozenClock

# -- Captured at collection time, i.e. before any test has run, so this is the canonical
# -- import-time registration state established by pptx/__init__.py.
_CANONICAL_PART_TYPES = dict(PartFactory.part_type_for)

_BIG_IO_MARK = "big_io"
_BIG_IO_SKIP_REASON = (
    "big_io: writes and reads a package larger than 256 MiB; enable with PAPER_BIG_IO=1 or "
    "`pytest -m big_io`"
)


def pytest_configure(config):
    # -- registered here (not in pyproject.toml) so no upstream config file changes; the
    # -- upstream suite runs warnings-as-errors and an unregistered mark would be an error.
    config.addinivalue_line(
        "markers",
        "lo_smoke: independent-loader smoke via headless LibreOffice; skipped when soffice "
        "is unavailable",
    )
    config.addinivalue_line(
        "markers",
        "big_io: package intake at a size that costs real disk I/O (>256 MiB written and "
        "read); skipped unless PAPER_BIG_IO=1 or selected with `-m big_io`",
    )


def pytest_collection_modifyitems(config, items):
    """Skip `big_io` tests unless the caller opted in, by env var or by `-m` selector."""
    if os.environ.get("PAPER_BIG_IO") or _BIG_IO_MARK in (config.getoption("-m") or ""):
        return
    skip = pytest.mark.skip(reason=_BIG_IO_SKIP_REASON)
    for item in items:
        if _BIG_IO_MARK in item.keywords:
            item.add_marker(skip)


@pytest.fixture(autouse=True)
def _canonical_part_factory_registrations():
    """Repair `PartFactory.part_type_for` before every paper test.

    Upstream's own unit tests mutate this class-level registry without restoring it
    (`DescribePartFactory` in tests/opc/test_package.py leaves a mock class registered for the
    slide-part content type), so in a combined run every presentation opened after that test
    loads its slide parts as MagicMocks. Paper tests are integration tests over real files and
    must always see the canonical import-time registrations, whatever ran before them.
    """
    PartFactory.part_type_for.clear()
    PartFactory.part_type_for.update(_CANONICAL_PART_TYPES)


@pytest.fixture
def frozen_clock():
    """A FrozenClock pinned to `tests.paper.clock.PAPER_TEST_INSTANT`."""
    return FrozenClock()
