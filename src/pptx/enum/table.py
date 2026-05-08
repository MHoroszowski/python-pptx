"""Built-in PowerPoint table style registry.

PowerPoint ships a fixed catalog of built-in table styles, each identified by a
GUID written into ``a:tbl/a:tblPr/a:tableStyleId``. The actual style
definitions live inside the PowerPoint application binary, not in the
``.pptx`` file — only the GUID reference is persisted, and PowerPoint
resolves it at render time.

This module exposes the GUIDs as a frozen mapping plus helper functions so
callers can apply a style by friendly name, e.g. ``"Medium Style 2 - Accent
1"``, instead of memorizing GUIDs. The registry covers the English-locale
names; PowerPoint stores the localized name in any saved
``ppt/tableStyles.xml`` it generates, but the GUID is locale-independent and
is what we round-trip.

Coverage notes:

- Phase 2 ships the most-cited subset harvested from real PowerPoint-authored
  decks (~39 entries), spanning the No Style / Themed / Light 1-3 / Medium
  1-3 / Dark 1-2 families.
- ``apply_style`` accepts either a friendly name (resolved against this
  registry) or a raw GUID string (passed through verbatim) — so any style not
  yet in the registry is still reachable via its GUID.
- Use ``register_table_style(name, guid)`` to extend the registry at runtime
  for styles not covered here (custom corp themes, additional Office
  built-ins discovered later).

References:

- ECMA-376 §21.1.3.15 (``CT_TableProperties``) — schema for ``tableStyleId``.
- The default ``ppt/tableStyles.xml`` shipped by PowerPoint contains only a
  single ``def`` GUID; the body of each built-in style is resolved
  internally. GUIDs in this file were harvested from multiple real
  PowerPoint-saved ``tableStyles.xml`` parts to ensure correctness.
"""

from __future__ import annotations

from typing import Mapping

# ---name -> GUID. GUIDs use canonical PowerPoint shape (brace-wrapped,
# ---upper-case hex). Keys are the English-locale `styleName` values
# ---PowerPoint writes into `tableStyles.xml`.
_BUILT_IN_TABLE_STYLES: Mapping[str, str] = {
    # No Style
    "No Style, No Grid": "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
    "No Style, Table Grid": "{5940675A-B579-460E-94D1-54222C63F5DA}",
    # Themed Styles
    "Themed Style 1 - Accent 1": "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}",
    "Themed Style 2 - Accent 1": "{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}",
    # Light Style 1 + accents
    "Light Style 1": "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}",
    "Light Style 1 - Accent 1": "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}",
    "Light Style 1 - Accent 2": "{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}",
    "Light Style 1 - Accent 3": "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}",
    "Light Style 1 - Accent 4": "{D27102A9-8310-4765-A935-A1911B00CA55}",
    "Light Style 1 - Accent 5": "{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}",
    "Light Style 1 - Accent 6": "{68D230F3-CF80-4859-8CE7-A43EE81993B5}",
    # Light Style 2 + accents
    "Light Style 2": "{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}",
    "Light Style 2 - Accent 1": "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}",
    "Light Style 2 - Accent 2": "{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}",
    "Light Style 2 - Accent 3": "{F2DE63D5-997A-4646-A377-4702673A728D}",
    "Light Style 2 - Accent 4": "{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}",
    "Light Style 2 - Accent 5": "{5A111915-BE36-4E01-A7E5-04B1672EAD32}",
    "Light Style 2 - Accent 6": "{912C8C85-51F0-491E-9774-3900AFEF0FD7}",
    # Light Style 3 + (partial) accents
    "Light Style 3": "{616DA210-FB5B-4158-B5E0-FEB733F419BA}",
    "Light Style 3 - Accent 1": "{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}",
    "Light Style 3 - Accent 6": "{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}",
    # Medium Style 1 + (partial) accents
    "Medium Style 1 - Accent 1": "{B301B821-A1FF-4177-AEE7-76D212191A09}",
    "Medium Style 1 - Accent 2": "{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}",
    "Medium Style 1 - Accent 3": "{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}",
    "Medium Style 1 - Accent 6": "{10A1B5D5-9B99-4C35-A422-299274C87663}",
    # Medium Style 2 + accents (the big one — Accent 1 is the python-pptx
    # default; Accents 1-6 are the most-cited entries in scanny#27)
    "Medium Style 2": "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}",
    "Medium Style 2 - Accent 1": "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
    "Medium Style 2 - Accent 2": "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}",
    "Medium Style 2 - Accent 3": "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}",
    "Medium Style 2 - Accent 4": "{00A15C55-8517-42AA-B614-E9B94910E393}",
    "Medium Style 2 - Accent 5": "{7DF18680-E054-41AD-8BC1-D1AEF772440D}",
    "Medium Style 2 - Accent 6": "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}",
    # Medium Style 3 + (partial) accents
    "Medium Style 3": "{8EC20E35-A176-4012-BC5E-935CFFF8708E}",
    "Medium Style 3 - Accent 4": "{EB9631B5-78F2-41C9-869B-9F39066F8104}",
    "Medium Style 3 - Accent 5": "{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}",
    # Dark Style 1/2
    "Dark Style 1": "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}",
    "Dark Style 2": "{5202B0CA-FC54-4496-8BCA-5EF66A818D29}",
    "Dark Style 2 - Accent 1/Accent 2": "{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}",
}


# ---PP_TABLE_STYLE is the public name. Exposed as a read-only mapping;
# ---callers extend via register_table_style(), not by mutating this dict.
PP_TABLE_STYLE: Mapping[str, str] = dict(_BUILT_IN_TABLE_STYLES)


# ---reverse-lookup table built once at import; mutable so register_table_style
# ---can keep both directions in sync.
_GUID_TO_NAME: dict[str, str] = {guid: name for name, guid in _BUILT_IN_TABLE_STYLES.items()}
# ---name lookup is case-insensitive: both directions stored lower-cased
_NAME_LOWER_TO_NAME: dict[str, str] = {name.lower(): name for name in _BUILT_IN_TABLE_STYLES}


def lookup_table_style(name: str) -> str:
    """Return the GUID for the built-in table style named `name`.

    Comparison is case-insensitive. Raises |ValueError| if `name` is not a
    known built-in style. Pass a raw GUID directly to ``Table.style_id`` /
    ``Table.apply_style`` instead of routing through this helper if your
    style isn't covered.
    """
    canonical = _NAME_LOWER_TO_NAME.get(name.lower())
    if canonical is None:
        raise ValueError("'%s' is not a known built-in table style name" % name)
    # ---refer to the live dict so register_table_style updates take effect
    return PP_TABLE_STYLE[canonical]  # pyright: ignore[reportIndexIssue]


def style_name_for(guid: str) -> str | None:
    """Return the friendly name registered for `guid`, or |None| when unknown.

    Comparison is exact (the registry stores GUIDs in canonical
    brace-wrapped upper-case-hex shape). Lossless fallback: callers can
    still use ``Table.style_id`` to read the raw GUID when no name is
    registered.
    """
    return _GUID_TO_NAME.get(guid)


def register_table_style(name: str, guid: str) -> None:
    """Add (or overwrite) a name → GUID entry in the public registry.

    Use this for built-in styles not yet covered by the shipped registry, or
    for custom ``tableStyles.xml`` entries embedded in your own templates.
    Both ``lookup_table_style(name)`` and ``style_name_for(guid)`` see the
    new entry immediately.
    """
    # ---PP_TABLE_STYLE is typed as read-only Mapping in public surface but
    # ---the underlying object is a dict; cast for the in-place update
    PP_TABLE_STYLE[name] = guid  # type: ignore[index]
    _GUID_TO_NAME[guid] = name
    _NAME_LOWER_TO_NAME[name.lower()] = name
