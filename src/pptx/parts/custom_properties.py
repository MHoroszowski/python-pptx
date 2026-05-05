"""Custom Document Properties part — `/docProps/custom.xml`."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.custom_properties import CT_Properties, CT_Property

if TYPE_CHECKING:
    from pptx.package import Package


class CustomPropertiesPart(XmlPart):
    """Corresponds to part named `/docProps/custom.xml`.

    Holds the package's custom (user-defined) document properties — the values
    surfaced under `File → Properties → Advanced` in PowerPoint. The
    user-facing Mapping wrapper lives at `pptx.custom_properties.CustomProperties`
    (Phase 3); this part just owns the XML and the per-property delegators.
    """

    _element: CT_Properties

    @classmethod
    def default(cls, package: "Package") -> "CustomPropertiesPart":
        """Return a new empty `CustomPropertiesPart` ready to add to `package`.

        Useful as the seed when a presentation doesn't yet have a custom
        properties part. The returned instance has no properties on it; the
        caller adds entries via `add_property(...)`.
        """
        return cls(
            PackURI("/docProps/custom.xml"),
            CT.OFC_CUSTOM_PROPERTIES,
            package,
            CT_Properties.new_properties(),
        )

    def add_property(self, name: str, value: object) -> CT_Property:
        """Add a new `<op:property>` for `(name, value)` and return it."""
        return self._element.add_property(name, value)

    def get_property(self, name: str) -> CT_Property | None:
        """Return the `<op:property>` with `name` or `None` if absent."""
        return self._element.get_property(name)

    def remove_property(self, name: str) -> bool:
        """Remove the `<op:property>` with `name`, returning True if found."""
        return self._element.remove_property(name)

    @property
    def property_names(self) -> tuple[str, ...]:
        """Tuple of property names in document order."""
        return self._element.property_names

    def __contains__(self, name: object) -> bool:
        return isinstance(name, str) and self._element.get_property(name) is not None

    def __iter__(self) -> Iterator[str]:
        return iter(self._element.property_names)

    def __len__(self) -> int:
        return len(self._element.property_lst)
