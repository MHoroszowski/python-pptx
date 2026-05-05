"""lxml custom element classes for the Custom Document Properties part.

Models `/docProps/custom.xml` — the `<op:Properties>` root and its `<op:property>`
children, each carrying one of five typed `<vt:*>` value elements.

Schema references: ECMA-376 Part 1, §15.2.12.2 (Custom File Properties Part).
"""

from __future__ import annotations

import datetime as dt
from typing import cast

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.simpletypes import XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

# Well-known FMTID Office writes on every user-defined custom property.
DEFAULT_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

# pid values 0 and 1 are reserved by the OOXML spec; user properties start at 2.
_FIRST_PID = 2

# Maximum string length for an lpwstr value. Office-tested limit; longer values
# round-trip but are reported by some inspectors as malformed.
_LPWSTR_MAX_LEN = 255


class CT_Properties(BaseOxmlElement):
    """`<op:Properties>` element, root of `/docProps/custom.xml`.

    The xmlchemy declaration is named `prop` rather than `property` because the
    latter would shadow Python's built-in `@property` decorator inside the
    class body — see metaclass-walk in `xmlchemy.py:120-131`. Public methods
    below preserve the `*_property` naming on the user-facing surface.
    """

    prop = ZeroOrMore("op:property", successors=())

    _properties_tmpl = "<op:Properties %s/>\n" % nsdecls("op", "vt")

    @staticmethod
    def new_properties() -> "CT_Properties":
        """Return a new empty `<op:Properties>` element with op + vt namespaces."""
        return cast("CT_Properties", parse_xml(CT_Properties._properties_tmpl))

    @property
    def property_lst(self) -> "list[CT_Property]":
        """List of `<op:property>` children in document order."""
        return cast("list[CT_Property]", self.prop_lst)

    def add_property(self, name: str, value: object) -> "CT_Property":
        """Append a new `<op:property>` child for `(name, value)`.

        The pid is auto-assigned to the next free integer ≥ 2 within this
        collection. Dispatches `value` by Python type to choose the `<vt:*>`
        child. Raises `TypeError` if `value` is not one of the supported types
        (see `CT_Property.value` for the dispatch table).
        """
        prop = cast("CT_Property", self._add_prop())
        prop.fmtid = DEFAULT_FMTID
        prop.pid = self._next_pid()
        prop.name = name
        prop.value = value
        return prop

    def get_property(self, name: str) -> "CT_Property | None":
        """Return the `<op:property>` child whose `name` attribute is `name`.

        Returns `None` if no such child exists. Match is case-sensitive — Office
        treats names case-sensitively even though Windows file names elsewhere
        do not.
        """
        for prop in self.property_lst:
            if prop.name == name:
                return prop
        return None

    def remove_property(self, name: str) -> bool:
        """Remove the `<op:property>` child with `name`, returning True if found."""
        prop = self.get_property(name)
        if prop is None:
            return False
        self.remove(prop)
        return True

    @property
    def property_names(self) -> tuple[str, ...]:
        """Tuple of `name` attributes for every `<op:property>` child, in order."""
        return tuple(p.name for p in self.property_lst)

    def _next_pid(self) -> int:
        """Return the next free pid (≥ 2) not yet used by any child."""
        used = {p.pid for p in self.property_lst if p.has_pid}
        candidate = _FIRST_PID
        while candidate in used:
            candidate += 1
        return candidate


class CT_Property(BaseOxmlElement):
    """`<op:property>` element — one custom document property entry."""

    fmtid: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "fmtid", XsdString
    )
    pid: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "pid", XsdUnsignedInt
    )
    name: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "name", XsdString
    )

    lpwstr = ZeroOrOne("vt:lpwstr", successors=())
    i4 = ZeroOrOne("vt:i4", successors=())
    r8 = ZeroOrOne("vt:r8", successors=())
    bool_ = ZeroOrOne("vt:bool", successors=())
    filetime = ZeroOrOne("vt:filetime", successors=())

    @property
    def has_pid(self) -> bool:
        """True if the `pid` attribute is present (it is required, but parsing
        a malformed file can leave it unset; this guards `_next_pid` against
        crashing on partial input)."""
        return self.get("pid") is not None

    @property
    def value(self) -> str | int | float | bool | dt.datetime | None:
        """The Python-typed value of whichever `<vt:*>` child is present.

        Returns `None` if no value child exists (a malformed but tolerated state).
        Order of precedence on read: lpwstr, i4, r8, bool, filetime — only one
        is expected to be present per the spec.
        """
        for child in (self.lpwstr, self.i4, self.r8, self.bool_, self.filetime):
            if child is not None:
                return cast("_VtValueElement", child).value_typed
        return None

    @value.setter
    def value(self, new_value: object) -> None:
        """Replace the current `<vt:*>` child with one matching `new_value`'s type.

        Dispatch table (bool checked BEFORE int because `bool` is a subclass of
        `int` in Python):

            bool                -> <vt:bool>
            int                 -> <vt:i4>
            float               -> <vt:r8>
            str                 -> <vt:lpwstr>
            datetime.datetime   -> <vt:filetime>

        Other types raise `TypeError`.
        """
        # Remove any existing value child before adding the new one.
        for tagname in ("vt:lpwstr", "vt:i4", "vt:r8", "vt:bool", "vt:filetime"):
            for elem in self.findall(qn(tagname)):
                self.remove(elem)

        if isinstance(new_value, bool):
            child = cast("CT_VtBool", self.get_or_add_bool_())
            child.value_typed = new_value
        elif isinstance(new_value, int):
            child = cast("CT_VtI4", self.get_or_add_i4())
            child.value_typed = new_value
        elif isinstance(new_value, float):
            child = cast("CT_VtR8", self.get_or_add_r8())
            child.value_typed = new_value
        elif isinstance(new_value, str):
            child = cast("CT_VtLpwstr", self.get_or_add_lpwstr())
            child.value_typed = new_value
        elif isinstance(new_value, dt.datetime):
            child = cast("CT_VtFiletime", self.get_or_add_filetime())
            child.value_typed = new_value
        else:
            raise TypeError(
                "custom property value must be bool, int, float, str, or datetime; "
                "got %s" % type(new_value).__name__
            )


class _VtValueElement(BaseOxmlElement):
    """Mixin-style base for `<vt:*>` typed value elements.

    Subclasses define a `value_typed` property that round-trips the element's
    text content to/from a Python value.
    """

    value_typed: object  # pyright: ignore[reportUninitializedInstanceVariable]


class CT_VtLpwstr(_VtValueElement):
    """`<vt:lpwstr>` — Unicode string value."""

    @property
    def value_typed(self) -> str:
        return self.text or ""

    @value_typed.setter
    def value_typed(self, value: str) -> None:
        if not isinstance(value, str):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError("vt:lpwstr value must be str, got %s" % type(value).__name__)
        if len(value) > _LPWSTR_MAX_LEN:
            raise ValueError(
                "vt:lpwstr value exceeds %d-character limit" % _LPWSTR_MAX_LEN
            )
        self.text = value


class CT_VtI4(_VtValueElement):
    """`<vt:i4>` — 32-bit signed integer value."""

    _MIN = -2147483648
    _MAX = 2147483647

    @property
    def value_typed(self) -> int:
        text = self.text
        if text is None:
            raise ValueError("vt:i4 element has no text content")
        return int(text)

    @value_typed.setter
    def value_typed(self, value: int) -> None:
        if isinstance(value, bool) or not isinstance(value, int):
            raise TypeError("vt:i4 value must be int, got %s" % type(value).__name__)
        if value < self._MIN or value > self._MAX:
            raise ValueError(
                "vt:i4 value out of range [%d, %d]: %d" % (self._MIN, self._MAX, value)
            )
        self.text = str(value)


class CT_VtR8(_VtValueElement):
    """`<vt:r8>` — IEEE-754 double-precision float value."""

    @property
    def value_typed(self) -> float:
        text = self.text
        if text is None:
            raise ValueError("vt:r8 element has no text content")
        return float(text)

    @value_typed.setter
    def value_typed(self, value: float) -> None:
        if isinstance(value, bool):
            raise TypeError("vt:r8 value must be float, got bool")
        if not isinstance(value, (int, float)):
            raise TypeError("vt:r8 value must be a number, got %s" % type(value).__name__)
        self.text = repr(float(value))


class CT_VtBool(_VtValueElement):
    """`<vt:bool>` — boolean value.

    Reads accept `"1"`, `"0"`, `"true"`, `"false"` (case-insensitive). Writes
    emit `"true"` or `"false"` to match what Microsoft Office produces.
    """

    @property
    def value_typed(self) -> bool:
        text = (self.text or "").strip().lower()
        if text in ("true", "1"):
            return True
        if text in ("false", "0"):
            return False
        raise ValueError("vt:bool element has invalid text content: %r" % self.text)

    @value_typed.setter
    def value_typed(self, value: bool) -> None:
        if not isinstance(value, bool):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError("vt:bool value must be bool, got %s" % type(value).__name__)
        self.text = "true" if value else "false"


class CT_VtFiletime(_VtValueElement):
    """`<vt:filetime>` — ISO-8601 UTC datetime value (always with `Z` suffix)."""

    @property
    def value_typed(self) -> dt.datetime:
        text = self.text
        if text is None:
            raise ValueError("vt:filetime element has no text content")
        return _parse_iso_utc(text)

    @value_typed.setter
    def value_typed(self, value: dt.datetime) -> None:
        if not isinstance(value, dt.datetime):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError(
                "vt:filetime value must be datetime, got %s" % type(value).__name__
            )
        # Office writes filetimes as UTC with a literal trailing 'Z'. If the
        # caller supplied a tz-aware value in another zone, convert; if naive,
        # assume already UTC (matches CorePropertiesPart's behavior).
        if value.tzinfo is not None:
            value = value.astimezone(dt.timezone.utc).replace(tzinfo=None)
        self.text = value.strftime("%Y-%m-%dT%H:%M:%SZ")


def _parse_iso_utc(text: str) -> dt.datetime:
    """Parse `text` as ISO-8601, returning a naive UTC `datetime`.

    Accepts the `Z` suffix Office writes and the `+HH:MM` form some tools use.
    Returns a naive datetime in UTC for symmetry with `_set_element_datetime`
    in `coreprops`. Raises `ValueError` on unparsable input.
    """
    cleaned = text.strip()
    if cleaned.endswith("Z"):
        cleaned = cleaned[:-1] + "+00:00"
    parsed = dt.datetime.fromisoformat(cleaned)
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone(dt.timezone.utc).replace(tzinfo=None)
    return parsed
