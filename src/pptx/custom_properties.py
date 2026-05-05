"""User-facing wrapper around the Custom Document Properties part.

Mapping-protocol surface that lets callers read and write the values exposed
under `File → Properties → Advanced` in PowerPoint as if they were a `dict`.
"""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, Iterator, Mapping, Union

if TYPE_CHECKING:
    from pptx.parts.custom_properties import CustomPropertiesPart


CustomPropertyValue = Union[str, int, float, bool, dt.datetime]


class CustomProperties(Mapping[str, CustomPropertyValue]):
    """Dict-like read/write access to custom document properties.

    Returned by :attr:`pptx.presentation.Presentation.custom_properties`. The
    mapping is *live* — writes go directly to the underlying
    `CustomPropertiesPart`; the next `Presentation.save(...)` persists them.

    Type dispatch on assignment is by Python type:

    ====================  ===================
    Python type           OOXML element
    ====================  ===================
    ``str``               ``<vt:lpwstr>``
    ``bool``              ``<vt:bool>``
    ``int``               ``<vt:i4>``
    ``float``             ``<vt:r8>``
    ``datetime.datetime`` ``<vt:filetime>``
    ====================  ===================

    For the cases where Python's type inference does the wrong thing — for
    example, you want a string `"42"` rather than the integer 42 — use the
    explicit :meth:`set_string` / :meth:`set_int` / etc. setters.
    """

    def __init__(self, part: "CustomPropertiesPart"):
        self._part = part

    # -- Mapping protocol --------------------------------------------------

    def __getitem__(self, name: str) -> CustomPropertyValue:
        prop = self._part.get_property(name)
        if prop is None:
            raise KeyError(name)
        value = prop.value
        if value is None:
            # Defensive: a malformed entry with no <vt:*> child is treated as
            # absent rather than surfacing None — keeps the Mapping contract clean.
            raise KeyError(name)
        return value

    def __setitem__(self, name: str, value: CustomPropertyValue) -> None:
        if not _is_supported(value):
            raise TypeError(
                "custom property value must be bool, int, float, str, or datetime; "
                "got %s" % type(value).__name__
            )
        existing = self._part.get_property(name)
        if existing is not None:
            existing.value = value
            return
        self._part.add_property(name, value)

    def __delitem__(self, name: str) -> None:
        if not self._part.remove_property(name):
            raise KeyError(name)

    def __contains__(self, name: object) -> bool:
        return isinstance(name, str) and self._part.get_property(name) is not None

    def __iter__(self) -> Iterator[str]:
        return iter(self._part.property_names)

    def __len__(self) -> int:
        return len(self._part)

    # -- Explicit-typed setters --------------------------------------------

    def set_string(self, name: str, value: str) -> None:
        """Write `value` as `<vt:lpwstr>` regardless of Python type."""
        if not isinstance(value, str):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError("set_string value must be str, got %s" % type(value).__name__)
        self._set_typed(name, value)

    def set_int(self, name: str, value: int) -> None:
        """Write `value` as `<vt:i4>` regardless of Python type.

        Rejects `bool` even though `bool` is-a `int` in Python — callers who
        really want a 1/0 i4 can wrap with `int(value)` first.
        """
        if isinstance(value, bool) or not isinstance(value, int):
            raise TypeError("set_int value must be int, got %s" % type(value).__name__)
        self._set_typed(name, value)

    def set_float(self, name: str, value: float) -> None:
        """Write `value` as `<vt:r8>` regardless of Python type."""
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise TypeError("set_float value must be a number, got %s" % type(value).__name__)
        self._set_typed(name, float(value))

    def set_bool(self, name: str, value: bool) -> None:
        """Write `value` as `<vt:bool>`."""
        if not isinstance(value, bool):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError("set_bool value must be bool, got %s" % type(value).__name__)
        self._set_typed(name, value)

    def set_datetime(self, name: str, value: dt.datetime) -> None:
        """Write `value` as `<vt:filetime>` (UTC, ISO-8601)."""
        if not isinstance(value, dt.datetime):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError("set_datetime value must be datetime, got %s" % type(value).__name__)
        self._set_typed(name, value)

    def _set_typed(self, name: str, value: CustomPropertyValue) -> None:
        """Replace-or-add the property; the underlying `CT_Property.value` setter
        already dispatches on Python type cleanly, so re-using it here is safe."""
        existing = self._part.get_property(name)
        if existing is not None:
            existing.value = value
            return
        self._part.add_property(name, value)


def _is_supported(value: object) -> bool:
    if isinstance(value, bool):
        return True
    return isinstance(value, (int, float, str, dt.datetime))
