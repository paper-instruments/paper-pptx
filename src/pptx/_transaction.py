"""Proxy-preserving rollback for paper-pptx package mutations."""

from __future__ import annotations

from contextvars import ContextVar
from typing import TYPE_CHECKING, Iterable

if TYPE_CHECKING:
    from pptx.opc.package import OpcPackage


_ACTIVE_TRANSACTIONS: ContextVar[tuple[object, ...]] = ContextVar(
    "paper_pptx_active_transactions", default=()
)


def package_has_open_transaction(package: "OpcPackage") -> bool:
    """Whether `package` has an unclosed transaction in the current context.

    Uses the same identity test as `PackageTransaction._is_outermost`, so a transaction
    on a *different* package never reports true. Contexts are per-thread and per-task, so
    a transaction open in another thread is correctly invisible here.
    """
    return any(
        transaction._package is package
        for transaction in _ACTIVE_TRANSACTIONS.get()
        if isinstance(transaction, PackageTransaction)
    )


class TransactionRollbackError(RuntimeError):
    """A mutation failed and one or more rollback steps also failed.

    This is deliberately not a ``PaperRefusal``. Once exact restoration cannot be
    established, callers must treat the package as unsafe rather than interpreting the
    operation as an atomic refusal.
    """

    def __init__(
        self,
        original_exception: BaseException,
        failures: Iterable[tuple[str, BaseException]],
    ):
        self.original_exception = original_exception
        self.failures = tuple(failures)
        details = "; ".join(
            "%s: %s(%s)" % (label, type(error).__name__, self._safe_message(error))
            for label, error in self.failures
        )
        super().__init__(
            "package rollback failed after %s: %d restore step(s) failed (%s)"
            % (
                type(original_exception).__name__,
                len(self.failures),
                details,
            )
        )

    @staticmethod
    def _safe_message(error: BaseException) -> str:
        """Return an exception message without trusting the exception's formatter."""
        try:
            return str(error)
        except BaseException as formatting_error:
            return "<unprintable; __str__ raised %s>" % type(formatting_error).__name__


class _ElementTreeState:
    """Original element objects and scalar state for one XML part."""

    def __init__(self, root):
        from lxml import etree

        self._root = root
        self._serialized = etree.tostring(root, encoding="UTF-8", standalone=True)
        self._original_prefixes = tuple(
            sorted(
                {
                    prefix
                    for element in root.iter()
                    for prefix in element.nsmap
                    if prefix is not None
                }
            )
        )
        self._records = tuple(
            (
                element,
                element.tag,
                dict(element.attrib),
                etree._Element.text.__get__(element),
                etree._Element.tail.__get__(element),
                tuple(element),
            )
            for element in root.iter()
        )

    def restore(self) -> None:
        """Restore the tree using the original nodes so existing proxies stay valid."""
        from lxml import etree

        original_elements = {record[0] for record in self._records}

        # Detach all current children first. This removes newly-created subtrees and
        # frees original nodes to be placed back under their original parents.
        for element, _, _, _, _, _ in self._records:
            for child in list(element):
                element.remove(child)
        for element in original_elements:
            if element is self._root:
                continue
            parent = element.getparent()
            if parent is not None:
                parent.remove(element)

        for element, tag, attributes, text, tail, _ in self._records:
            if isinstance(tag, str) and element.tag != tag:
                element.tag = tag
            if dict(element.attrib) != attributes:
                element.attrib.clear()
                element.attrib.update(attributes)
            # Some oxml classes override `.text` with a domain property lacking a
            # setter. Address the underlying lxml descriptors directly.
            if etree._Element.text.__get__(element) != text:
                etree._Element.text.__set__(element, text)
            if etree._Element.tail.__get__(element) != tail:
                etree._Element.tail.__set__(element, tail)
        for element, _, _, _, _, children in self._records:
            for child in children:
                element.append(child)

        # lxml retains namespace declarations introduced by a removed namespaced
        # attribute. Remove only declarations that are now unused, preserving every
        # prefix that existed before the transaction. Exact verification below remains
        # authoritative for declarations cleanup cannot restore in place.
        if self.current_serialized != self._serialized:
            etree.cleanup_namespaces(
                self._root, keep_ns_prefixes=self._original_prefixes
            )

    @property
    def current_serialized(self) -> bytes:
        """Current part serialization in the same form used by ``XmlPart.blob``."""
        from lxml import etree

        return etree.tostring(self._root, encoding="UTF-8", standalone=True)

    @property
    def is_exact(self) -> bool:
        """Whether serialization, scalar state, children, and original nodes all match."""
        if self.current_serialized != self._serialized:
            return False
        from lxml import etree

        return all(
            element.tag == tag
            and dict(element.attrib) == attributes
            and etree._Element.text.__get__(element) == text
            and etree._Element.tail.__get__(element) == tail
            and tuple(element) == children
            for element, tag, attributes, text, tail, children in self._records
        )


class _ObjectGraphState:
    """Mutable state reachable through existing package and caller proxy caches."""

    def __init__(self, roots: Iterable[object]):
        from lxml import etree

        from pptx.opc.package import OpcPackage, Part

        self._instance_states: list[tuple[object, dict]] = []
        self._container_states: list[tuple[object, object]] = []
        seen: set[int] = set()
        stack = list(roots)

        while stack:
            value = stack.pop()
            value_id = id(value)
            if value_id in seen:
                continue
            seen.add(value_id)

            if isinstance(value, (str, bytes, int, float, bool, type(None))):
                continue
            if isinstance(value, (OpcPackage, Part, etree._Element)):
                continue
            if isinstance(value, bytearray):
                self._container_states.append((value, bytes(value)))
                continue
            if isinstance(value, dict):
                before = dict(value)
                self._container_states.append((value, before))
                stack.extend(before.keys())
                stack.extend(before.values())
                continue
            if isinstance(value, list):
                before = tuple(value)
                self._container_states.append((value, before))
                stack.extend(before)
                continue
            if isinstance(value, set):
                before = set(value)
                self._container_states.append((value, before))
                stack.extend(before)
                continue
            if isinstance(value, (tuple, frozenset)):
                stack.extend(value)
                continue

            # Proxy and collaborator objects in this package use ``__dict__`` for
            # lazy-property and performance caches. Avoid walking arbitrary user objects.
            if not type(value).__module__.startswith("pptx.") or not hasattr(value, "__dict__"):
                continue
            before = dict(value.__dict__)
            self._instance_states.append((value, before))
            stack.extend(before.values())

    def restore(self) -> list[tuple[str, BaseException]]:
        """Restore every cached instance and container, returning all failures."""
        failures: list[tuple[str, BaseException]] = []
        for instance, before in self._instance_states:
            try:
                PackageTransaction._restore_instance_dict(instance, before)
            except BaseException as error:
                failures.append((self._label("object", instance), error))
        for container, before in self._container_states:
            try:
                self._restore_container(container, before)
            except BaseException as error:
                failures.append((self._label("container", container), error))
        return failures

    @staticmethod
    def _label(kind: str, value: object) -> str:
        """Return a stable diagnostic label for one in-memory restore target."""
        cls = type(value)
        return "%s %s.%s@0x%x" % (kind, cls.__module__, cls.__qualname__, id(value))

    @staticmethod
    def _restore_container(container: object, before: object) -> None:
        """Restore one supported mutable container without replacing its identity."""
        if isinstance(container, dict):
            container.clear()
            container.update(before)
        elif isinstance(container, (bytearray, list)):
            container[:] = before
        else:
            container.clear()
            container.update(before)


class _ValidationReadState:
    """Live structural state restored after serialization-only reads.

    Save preflight can materialize lazy caches, and a custom ``XmlPart.blob`` getter can mutate
    its live element. Capturing both object state and original element nodes allows those read-side
    effects to be removed without undoing edits made before this snapshot.
    """

    def __init__(
        self,
        transaction: "PackageTransaction",
        roots: Iterable[object] = (),
        *,
        snapshot_all_xml: bool = True,
    ):
        from pptx.opc.package import XmlPart

        self._transaction = transaction
        package = transaction._package
        self._package_dict = dict(package.__dict__)
        self._part_states = []
        graph_roots = list(self._package_dict.values())
        for part, part_dict in transaction._validated_reachable_part_dicts(
            self._package_dict
        ):
            element_state = (
                _ElementTreeState(part._element)
                if isinstance(part, XmlPart)
                and (snapshot_all_xml or type(part).blob is not XmlPart.blob)
                else None
            )
            part_label = str(part_dict.get("_partname", type(part).__name__))
            self._part_states.append((part, part_dict, element_state, part_label))
            graph_roots.extend(part_dict.values())
        graph_roots.extend(roots)
        self._object_graph_state = _ObjectGraphState(graph_roots)

    def restore(
        self, *, force_element_restore: bool = True
    ) -> list[tuple[str, BaseException]]:
        """Restore captured live state, returning every cleanup failure."""
        failures: list[tuple[str, BaseException]] = []
        failed_xml_state_ids: set[int] = set()
        for _, _, element_state, part_label in self._part_states:
            if element_state is None:
                continue
            try:
                if force_element_restore or not element_state.is_exact:
                    element_state.restore()
            except BaseException as error:
                failed_xml_state_ids.add(id(element_state))
                failures.append(("XML part %s" % part_label, error))

        for part, part_dict, _, part_label in self._part_states:
            try:
                PackageTransaction._restore_instance_dict(part, part_dict)
            except BaseException as error:
                failures.append(("part %s" % part_label, error))
        try:
            PackageTransaction._restore_instance_dict(
                self._transaction._package, self._package_dict
            )
        except BaseException as error:
            failures.append(("package", error))
        failures.extend(self._object_graph_state.restore())

        for _, _, element_state, part_label in self._part_states:
            if element_state is None or id(element_state) in failed_xml_state_ids:
                continue
            try:
                if not element_state.is_exact:
                    raise RuntimeError("XML serialization or node identity differs after restore")
            except BaseException as error:
                failures.append(("XML part %s" % part_label, error))
        return failures


class PackageTransaction:
    """Restore a package's original live graph if the guarded operation raises."""

    def __init__(self, package: "OpcPackage", *roots: object):
        self._package = package
        self._roots = tuple(roots)
        self._package_dict = {}
        self._part_states = []
        self._object_graph_state = None
        self._context_token = None
        self._is_outermost = False

        # Preserve constructor-time refusal for malformed relationships and signatures without
        # freezing rollback state before the context is actually entered.
        self._validated_reachable_part_dicts(dict(package.__dict__))

    def __enter__(self) -> "PackageTransaction":
        if self._context_token is not None:
            raise RuntimeError("PackageTransaction cannot be entered more than once")
        self._capture_entry_snapshot()
        active = _ACTIVE_TRANSACTIONS.get()
        self._is_outermost = not any(
            transaction._package is self._package
            for transaction in active
            if isinstance(transaction, PackageTransaction)
        )
        self._context_token = _ACTIVE_TRANSACTIONS.set(active + (self,))
        return self

    def _capture_entry_snapshot(self) -> None:
        """Capture rollback state at context entry and reverse blob-getter read effects."""
        from pptx.opc.package import XmlPart

        state = _ValidationReadState(self, self._roots)
        payloads = []
        try:
            for part, _, element_state, _ in state._part_states:
                payloads.append(
                    element_state._serialized
                    if element_state is not None and type(part).blob is XmlPart.blob
                    else bytes(part.blob)
                )
        except BaseException as error:
            cleanup_failures = state.restore(force_element_restore=False)
            if cleanup_failures:
                raise TransactionRollbackError(error, cleanup_failures) from error
            raise

        cleanup_failures = state.restore(force_element_restore=False)
        if cleanup_failures:
            snapshot_error = RuntimeError(
                "transaction snapshot serialization changed live state and cleanup failed"
            )
            raise TransactionRollbackError(snapshot_error, cleanup_failures) from snapshot_error

        self._package_dict = state._package_dict
        self._part_states = [
            (part, part_dict, element_state, payload, part_label)
            for (part, part_dict, element_state, part_label), payload in zip(
                state._part_states, payloads
            )
        ]
        self._object_graph_state = state._object_graph_state

    def __exit__(self, exc_type, exc_value, traceback) -> bool:
        token = self._context_token
        if token is None:
            raise RuntimeError("PackageTransaction exited without being entered")
        active = _ACTIVE_TRANSACTIONS.get()
        if not active or active[-1] is not self:
            raise RuntimeError("PackageTransaction contexts exited out of order")
        try:
            if exc_type is not None:
                assert exc_value is not None
                self._rollback(exc_value)
                return False
            if not self._is_outermost:
                return False
            try:
                self._validate_candidate()
            except BaseException as validation_error:
                self._rollback(validation_error)
                raise
            return False
        finally:
            _ACTIVE_TRANSACTIONS.reset(token)
            self._context_token = None

    def _rollback(self, original_exception: BaseException) -> None:
        """Restore the pre-transaction graph or raise a loud rollback failure."""
        assert self._object_graph_state is not None
        failures: list[tuple[str, BaseException]] = []
        failed_xml_state_ids: set[int] = set()

        # Restore every independent state even when an earlier target is hostile. A
        # partial best effort is still valuable, but it is never presented as atomic.
        for _, _, element_state, _, part_label in self._part_states:
            if element_state is None:
                continue
            try:
                element_state.restore()
            except BaseException as error:
                failed_xml_state_ids.add(id(element_state))
                failures.append(("XML part %s" % part_label, error))

        for part, part_dict, _, _, part_label in self._part_states:
            try:
                self._restore_instance_dict(part, part_dict)
            except BaseException as error:
                failures.append(("part %s" % part_label, error))

        try:
            self._restore_instance_dict(self._package, self._package_dict)
        except BaseException as error:
            failures.append(("package", error))

        failures.extend(self._object_graph_state.restore())
        for part, _, element_state, expected_payload, part_label in self._part_states:
            if element_state is not None:
                if id(element_state) in failed_xml_state_ids:
                    continue
                try:
                    if not element_state.is_exact:
                        raise RuntimeError("XML serialization differs after rollback")
                    restored_payload = bytes(part.blob)
                    if restored_payload != expected_payload:
                        raise RuntimeError(
                            "XML payload differs after rollback (%d bytes before, %d after)"
                            % (len(expected_payload), len(restored_payload))
                        )
                except BaseException as error:
                    failures.append(("XML part %s" % part_label, error))
                finally:
                    # A custom blob getter can itself mutate the live element. Restore the
                    # original nodes again after observing its save-visible payload.
                    try:
                        element_state.restore()
                        if not element_state.is_exact:
                            raise RuntimeError(
                                "XML serialization or node identity differs after payload "
                                "verification"
                            )
                    except BaseException as error:
                        failures.append(("XML part %s" % part_label, error))
                continue
            try:
                restored_payload = bytes(part.blob)
                if restored_payload != expected_payload:
                    raise RuntimeError(
                        "binary payload differs after rollback (%d bytes before, %d after)"
                        % (len(expected_payload), len(restored_payload))
                    )
            except BaseException as error:
                failures.append(("binary part %s" % part_label, error))

        # Payload getters can also materialize caches or mutate supported collaborators. Remove
        # those verification-only effects before deciding whether rollback was exact.
        for part, part_dict, _, _, part_label in self._part_states:
            try:
                self._restore_instance_dict(part, part_dict)
            except BaseException as error:
                failures.append(("part %s after payload verification" % part_label, error))
        try:
            self._restore_instance_dict(self._package, self._package_dict)
        except BaseException as error:
            failures.append(("package after payload verification", error))
        failures.extend(self._object_graph_state.restore())
        if failures:
            raise TransactionRollbackError(original_exception, failures) from original_exception

    def _validate_candidate(self) -> None:
        """Validate the complete candidate using the ordinary-save output contract."""
        validation_state = _ValidationReadState(self)
        self._stage_and_reopen_candidate()
        cleanup_failures = validation_state.restore(force_element_restore=False)
        if cleanup_failures:
            raise RuntimeError(
                "candidate validation changed live package state and cleanup failed (%s)"
                % "; ".join(
                    "%s: %s" % (label, type(error).__name__)
                    for label, error in cleanup_failures
                )
            )

    def _stage_and_reopen_candidate(self) -> None:
        """Serialize privately and validate through the public presentation loader."""
        from io import BytesIO

        from pptx.api import Presentation
        from pptx.errors import UnsupportedStructureError

        candidate = BytesIO()
        self._package.save(candidate)
        candidate.seek(0)
        try:
            Presentation(candidate)
        except Exception as exc:
            raise UnsupportedStructureError(
                "transaction candidate is not a reopenable PowerPoint presentation (%s)" % exc
            ) from exc

    def _validated_reachable_part_dicts(
        self, package_dict: dict
    ) -> list[tuple[object, dict]]:
        """Return reachable part states after raw graph and signature validation."""
        from pptx.errors import UnsupportedStructureError
        from pptx.opc.constants import CONTENT_TYPE as CT

        signature_content_types = {
            CT.OPC_DIGITAL_SIGNATURE_CERTIFICATE,
            CT.OPC_DIGITAL_SIGNATURE_ORIGIN,
            CT.OPC_DIGITAL_SIGNATURE_XMLSIGNATURE,
        }
        part_states = []
        seen_part_ids: set[int] = set()
        pending = list(self._internal_targets(package_dict))
        while pending:
            part = pending.pop()
            part_id = id(part)
            if part_id in seen_part_ids:
                continue
            seen_part_ids.add(part_id)
            part_dict = dict(part.__dict__)
            if part_dict.get("_content_type") in signature_content_types:
                raise UnsupportedStructureError(
                    "mutating a digitally signed presentation would invalidate its OPC "
                    "signature; remove or re-sign it with a signature-aware tool first"
                )
            part_states.append((part, part_dict))
            pending.extend(self._internal_targets(part_dict))
        return part_states

    def _internal_targets(self, owner_state: dict) -> list[object]:
        """Return internal target parts without evaluating any lazy properties."""
        from pptx.errors import RelationshipPolicyError, UnsupportedStructureError
        from pptx.opc.constants import RELATIONSHIP_TARGET_MODE as RTM
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        from pptx.opc.package import Part

        targets = []
        seen_relationship_collections: set[int] = set()
        for key in ("_rels", "rels"):
            relationships = owner_state.get(key)
            if relationships is None or id(relationships) in seen_relationship_collections:
                continue
            seen_relationship_collections.add(id(relationships))
            relationship_map = getattr(relationships, "__dict__", {}).get("_rels")
            if relationship_map is None:
                continue
            if not isinstance(relationship_map, dict):
                raise RelationshipPolicyError(
                    "relationship collection on %s is malformed"
                    % owner_state.get("_partname", "/")
                )
            for relationship in relationship_map.values():
                relationship_state = getattr(relationship, "__dict__", {})
                if relationship_state.get("_reltype") in {RT.ORIGIN, RT.SIGNATURE}:
                    raise UnsupportedStructureError(
                        "mutating a digitally signed presentation would invalidate its OPC "
                        "signature; remove or re-sign it with a signature-aware tool first"
                    )
                target_mode = relationship_state.get("_target_mode")
                if target_mode == RTM.EXTERNAL:
                    continue
                if target_mode != RTM.INTERNAL:
                    raise RelationshipPolicyError(
                        "relationship %s has invalid target mode %r"
                        % (relationship_state.get("_rId", "<unknown>"), target_mode)
                    )
                target = relationship_state.get("_target")
                if not isinstance(target, Part):
                    raise RelationshipPolicyError(
                        "internal relationship %s targets %s, not a package Part"
                        % (
                            relationship_state.get("_rId", "<unknown>"),
                            type(target).__name__,
                        )
                    )
                target_state = target.__dict__
                if target_state.get("_package") is not self._package:
                    raise RelationshipPolicyError(
                        "internal relationship %s targets part %s owned by another package"
                        % (
                            relationship_state.get("_rId", "<unknown>"),
                            target_state.get("_partname", "<unknown>"),
                        )
                    )
                targets.append(target)
        return targets

    @staticmethod
    def _restore_instance_dict(instance, before: dict) -> None:
        """Restore the exact shallow instance state, including all lazy-cache values."""
        current = instance.__dict__
        current.clear()
        current.update(before)
