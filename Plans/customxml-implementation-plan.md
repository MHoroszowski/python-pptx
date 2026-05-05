# Plan: customXml part manipulation in `python-pptx-extended`

> **Status:** proposal — awaiting principal approval before implementation begins.
> **Scope:** add first-class read/write support for the two OOXML mechanisms that
> let an application embed structured data in a `.pptx`:
>
> 1. **Custom document properties** — `/docProps/custom.xml` (visible in PowerPoint UI under *File → Properties → Advanced*).
> 2. **CustomXml data parts** — `/customXml/itemN.xml` + `/customXml/itemPropsN.xml` (hidden from end users; the mechanism Office.js, SharePoint, and VSTO use).
>
> The first consumer is a CLI that round-trips a markdown source document, but the
> public API is general-purpose: provenance metadata, AI generation markers,
> template parameters, application-specific configuration, etc.

---

## 1. Context

### Why this fork

Mainline `scanny/python-pptx` v0.4.1 made the loader *tolerate* customXml parts
(parts no longer trip the importer when present), but never exposed an API to
read, mutate, or create them. Issues
[#286](https://github.com/scanny/python-pptx/issues/286) (custom doc properties)
and [#578](https://github.com/scanny/python-pptx/issues/578) (custom tags) have
been open and unaddressed for years. Other forks (`python-pptx-ng`,
`python-pptx-fix`, `python-pptx-fork`) inherit the same gap.

The pattern we are porting comes from
[`python-openxml/python-docx-oss`](https://github.com/python-openxml/python-docx-oss),
which solved the equivalent problem for `.docx` (`document.custom_properties`,
`document.part.custom_xml_parts[i].add_item(...)`). We adapt that surface to
PresentationML's relationship topology — most importantly, customXml data parts
must hang off `ppt/presentation.xml.rels` (presentation-scoped), not the package
root, or Office.js will not enumerate them
([MS Q&A](https://learn.microsoft.com/en-us/answers/questions/5586825/how-to-add-a-proper-customxml-to-a-powerpoint-pres)).

### What the existing code already gives us for free

A short codebase survey before signature design saved a lot of plumbing work:

| Concern | Already in the fork | Where |
|---|---|---|
| Content-type constants | `CT.OFC_CUSTOM_PROPERTIES`, `CT.OFC_CUSTOM_XML_PROPERTIES`, `CT.XML` | `src/pptx/opc/constants.py:33–34, 170` |
| Relationship-type constants | `RT.CUSTOM_PROPERTIES`, `RT.CUSTOM_XML`, `RT.CUSTOM_XML_PROPS` | `src/pptx/opc/constants.py:220–229` |
| Auto-derived `[Content_Types].xml` | `_ContentTypesItem._defaults_and_overrides` reads `part.content_type` for every part; `xml` extension defaults to `application/xml` so `customXml/itemN.xml` lands under the default with no extra wiring | `src/pptx/opc/serialized.py:280–296` |
| Part-class registration | `PartFactory.part_type_for.update({...})` at module load | `src/pptx/__init__.py:35–69` |
| Pattern for property-style XML parts | `CorePropertiesPart` + `CT_CoreProperties` — a sibling pair we can copy | `src/pptx/parts/coreprops.py`, `src/pptx/oxml/coreprops.py` |
| Package-root vs. presentation-scoped relating | `package.relate_to(part, RT.X)` writes `/_rels/.rels`; `presentation_part.relate_to(part, RT.X)` writes `/ppt/_rels/presentation.xml.rels` | `src/pptx/opc/package.py:41–51, 357–361` |
| Lazy-load with graceful re-use | `lazyproperty` + `try part_related_by(...) / except KeyError: create-and-relate` | `src/pptx/package.py:19–30` (CoreProperties pattern) |
| `xmlchemy` machinery | `BaseOxmlElement`, `ZeroOrOne`, `ZeroOrMore`, `OptionalAttribute`, `RequiredAttribute`, `register_element_cls` | `src/pptx/oxml/xmlchemy.py`, `src/pptx/oxml/__init__.py` |

**So no changes to constants, content-type registration, or the package writer
are required.** The work is: add new oxml classes, two new part subclasses, two
new collection wrappers, hang two properties off `Presentation`, and register
two content-types in `__init__.py`.

---

## 2. Public API design

> All examples assume `prs = Presentation("input.pptx")`. `Presentation` is the
> existing `pptx.presentation.Presentation` class.

### 2.1 `Presentation.custom_properties` — typed dict-like

Mirrors the docx-oss `CustomProperties` API. Each property is a `<property>`
element under `/docProps/custom.xml`; values are typed via the `vt:` namespace
(`lpwstr`, `i4`, `r8`, `bool`, `filetime`).

```python
class CustomProperties(Mapping[str, "CustomPropertyValue"]):
    """Read/write Custom document properties (visible in PowerPoint UI)."""

    def __getitem__(self, name: str) -> str | int | float | bool | datetime: ...
    def __setitem__(self, name: str, value: str | int | float | bool | datetime) -> None: ...
    def __delitem__(self, name: str) -> None: ...
    def __contains__(self, name: object) -> bool: ...
    def __iter__(self) -> Iterator[str]: ...
    def __len__(self) -> int: ...

    def get(self, name: str, default=None): ...
    def keys(self) -> KeysView[str]: ...
    def items(self) -> ItemsView[str, "CustomPropertyValue"]: ...
    def values(self) -> ValuesView["CustomPropertyValue"]: ...

    # Explicit-typed setters when the dispatch by Python type is wrong
    def set_string(self, name: str, value: str) -> None: ...
    def set_int(self, name: str, value: int) -> None: ...
    def set_float(self, name: str, value: float) -> None: ...
    def set_bool(self, name: str, value: bool) -> None: ...
    def set_datetime(self, name: str, value: datetime) -> None: ...
```

```python
prs = Presentation("input.pptx")
prs.custom_properties["Source"] = "deck-builder-cli@1.4.2"
prs.custom_properties["GeneratedAt"] = datetime.now(timezone.utc)
prs.custom_properties["BuildNumber"] = 42
prs.custom_properties.set_string("FreeformNotes", "anything goes here")
del prs.custom_properties["Stale"]
prs.save("output.pptx")
```

**Type dispatch by Python type at `__setitem__`:**

| Python type | `vt:` element |
|---|---|
| `str` | `vt:lpwstr` |
| `bool` (checked **before** `int`) | `vt:bool` |
| `int` | `vt:i4` |
| `float` | `vt:r8` |
| `datetime.datetime` | `vt:filetime` |

Anything else raises `TypeError`. The explicit `set_*` methods exist for the
case where the caller wants `lpwstr` *string* representations of numbers, or
where future types are added (`vt:lpstr`, `vt:r8`, etc.).

**`fmtid` and `pid`:** every `<property>` element requires
`fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"` (the well-known Office FMTID)
and a `pid` ≥ 2 unique within the part. The collection auto-assigns `pid` (next
free integer ≥ 2). Callers never see `pid`.

### 2.2 `Presentation.custom_xml_parts` — collection of arbitrary-XML parts

Mirrors docx-oss `document.part.custom_xml_parts`. Each entry is a
`CustomXmlPart` paired with a `CustomXmlPropertiesPart` (its `itemPropsN.xml`
sibling carrying the `datastoreItem` GUID and any `schemaRefs`).

```python
class CustomXmlParts(Sequence["CustomXmlPart"]):
    """Collection of customXml data parts attached to the presentation."""

    def __getitem__(self, key: int | str) -> "CustomXmlPart":
        """Index by integer position OR by part name (e.g. 'item3.xml').

        Use `.by_guid(...)` for datastoreItem-id lookup.
        """

    def __iter__(self) -> Iterator["CustomXmlPart"]: ...
    def __len__(self) -> int: ...

    def by_guid(self, guid: str) -> "CustomXmlPart | None":
        """Lookup by datastoreItem id (the GUID in itemPropsN.xml). Match is
        case-insensitive and curly-brace-tolerant."""

    def by_name(self, name: str) -> "CustomXmlPart | None":
        """Lookup by application-assigned name. Names live in custom_properties
        under a reserved `_pptx_customxml_name_<id>` key — see §3.4."""

    def add(
        self,
        xml: bytes | str | "lxml.etree._Element",
        *,
        name: str | None = None,
        datastoreItem_id: str | None = None,
        schema_refs: Iterable[str] | None = None,
        scope: Literal["presentation", "package"] = "presentation",
    ) -> "CustomXmlPart":
        """Add a new customXml part with the given XML payload.

        Parameters
        ----------
        xml
            Raw XML — bytes, str, or an lxml `_Element`. Must be well-formed
            XML; the caller owns the root element name and namespaces. Stored
            verbatim (modulo any normalization lxml does on parse).
        name
            Optional application-assigned name. Stored as a custom document
            property under `_pptx_customxml_name_<datastoreItem_id>`. See §3.4
            for why we do not use the `<name>` attribute on `customXmlPart`.
        datastoreItem_id
            Optional GUID. If omitted, a new `uuid4()` is generated and wrapped
            in curly braces ("{...}") to match Office's format.
        schema_refs
            Optional iterable of schema namespace URIs that this customXml part
            claims to conform to. Written as `<ds:schemaRef ds:uri="...">`
            children of `<ds:datastoreItem>` in itemProps.
        scope
            "presentation" (default) writes the relationship into
            `ppt/_rels/presentation.xml.rels` — the topology Office.js
            enumerates. "package" writes to `_rels/.rels` to match VSTO /
            SharePoint patterns. The two are not exchangeable round-trip
            (PowerPoint preserves the topology it was written with).

        Returns
        -------
        The new `CustomXmlPart`. Already attached; nothing else to do before
        `prs.save()`.
        """

    def remove(self, part: "CustomXmlPart | int | str") -> None:
        """Remove the part (and its paired CustomXmlPropertiesPart) from the
        package. Drops the relationship from whichever source (presentation or
        package) currently owns it. Idempotent if already removed."""
```

```python
class CustomXmlPart:
    """A single customXml/itemN.xml + customXml/itemPropsN.xml pair."""

    @property
    def name(self) -> str | None:
        """Application-assigned name from custom_properties, or None."""

    @property
    def datastoreItem_id(self) -> str:
        """GUID identifying the part across edits (e.g. '{1A2B...}')."""

    @datastoreItem_id.setter
    def datastoreItem_id(self, value: str) -> None: ...

    @property
    def schema_refs(self) -> tuple[str, ...]:
        """Tuple of `ds:schemaRef ds:uri` values from itemProps."""

    @schema_refs.setter
    def schema_refs(self, value: Iterable[str]) -> None: ...

    @property
    def scope(self) -> Literal["presentation", "package"]:
        """Where this part's relationship is currently rooted (read-only;
        change via remove + re-add)."""

    @property
    def partname(self) -> str:
        """Package URI, e.g. '/customXml/item3.xml'."""

    @property
    def element(self) -> "lxml.etree._Element":
        """Live root element of the customXml payload. Mutating it mutates the
        part. For replace-whole-payload semantics, use `.replace_xml(...)`."""

    @property
    def blob(self) -> bytes:
        """Serialized bytes of the customXml payload (with XML declaration)."""

    def replace_xml(self, xml: bytes | str | "lxml.etree._Element") -> None:
        """Replace the entire payload with `xml`. The root element is
        replaced, not merged. Preserves datastoreItem_id and schema_refs (those
        live in the sibling itemProps part)."""

    # docx-oss compatibility shim — only present if we adopt it (see §8 Q1)
    def add_item(self, tag: str, text: str = "", **attrs: str) -> "lxml.etree._Element":
        """Append a child element `<tag>text</tag>` with attributes. Returns the
        appended element. Convenience for the common "flat list of items"
        shape; for arbitrary structure use `.element` directly."""
```

```python
import uuid
prs = Presentation("input.pptx")

# General case — arbitrary XML
part = prs.custom_xml_parts.add(
    b"""<?xml version="1.0"?>
        <provenance xmlns="urn:my-app:provenance">
          <source>deck-builder-cli</source>
          <built-at>2026-05-05T14:00:00Z</built-at>
        </provenance>""",
    name="provenance",
    schema_refs=["urn:my-app:provenance"],
)
print(part.datastoreItem_id)  # auto-assigned GUID

# Lookup
same = prs.custom_xml_parts.by_name("provenance")
assert same is part
also_same = prs.custom_xml_parts.by_guid(part.datastoreItem_id)

# Mutate
same.element.find("{urn:my-app:provenance}source").text = "deck-builder-cli@1.4.3"

prs.save("output.pptx")
```

### 2.3 String-blob helper — the primary use case

Most callers want "stash this string verbatim, give it back to me on read."
Wrapping it in a one-element XML envelope keeps it valid OOXML and lets the
mime hint round-trip:

```python
def add_string_blob(
    self,
    name: str,
    content: str,
    *,
    mime_hint: str | None = None,
    encoding: Literal["text", "base64"] = "text",
    scope: Literal["presentation", "package"] = "presentation",
) -> "CustomXmlPart":
    """Embed a string payload as a customXml part.

    Wraps `content` in:
        <blob xmlns="urn:python-pptx:blob"
              name="..." mime="..." encoding="text|base64">...</blob>

    For binary or non-XML-safe text, set encoding="base64" and pass already-
    encoded content (the helper does NOT auto-base64; the caller is
    responsible). Round-trip: read with `.element.text` or via the helper
    `read_string_blob(name)`."""

def read_string_blob(self, name: str) -> str | None:
    """Return content of the blob part with `name`, or None if not present.
    If encoding='base64', returns the still-encoded string — the caller
    decodes."""
```

The `urn:python-pptx:blob` envelope namespace is reserved for this fork's
helpers. Callers using `.add(...)` directly are free to use any namespace they
want.

### 2.4 Property accessors on `Presentation`

Two new properties on `pptx.presentation.Presentation`:

```python
@property
def custom_properties(self) -> CustomProperties:
    """CustomProperties instance for /docProps/custom.xml. Created on first
    access if the part does not yet exist (consistent with .core_properties).
    """

@property
def custom_xml_parts(self) -> CustomXmlParts:
    """Collection of customXml data parts. Always returns the same collection
    instance for a given Presentation."""
```

Both delegate through `self.part` to the `PresentationPart`, which owns the
lazy-loaded helpers — same pattern as `core_properties`.

---

## 3. Internal architecture

### 3.1 New files

| Path | Purpose |
|---|---|
| `src/pptx/parts/custom_properties.py` | `CustomPropertiesPart(XmlPart)` — `/docProps/custom.xml` |
| `src/pptx/parts/custom_xml.py` | `CustomXmlPart(XmlPart)`, `CustomXmlPropertiesPart(XmlPart)` |
| `src/pptx/oxml/custom_properties.py` | `CT_CustomProperties`, `CT_Property`, value-type element classes |
| `src/pptx/oxml/custom_xml.py` | `CT_DatastoreItem`, `CT_DatastoreSchemaRef` |
| `src/pptx/custom_properties.py` | `CustomProperties` (Mapping wrapper) |
| `src/pptx/custom_xml.py` | `CustomXmlParts` (Sequence wrapper), `CustomXmlPart` user-facing facade |

Layering rationale (matches the rest of the codebase):

- `oxml/*` — pure XML element classes; no relationship logic; xmlchemy types only.
- `parts/*` — `XmlPart` subclasses; own a single `_element`; `lazyproperty`
  helpers but no end-user collections.
- `custom_properties.py`, `custom_xml.py` (top-level) — user-facing wrappers
  (Mapping/Sequence) that the principal hangs off `Presentation`. Mirrors how
  `pptx/slide.py` (`Slides`, `SlideMasters`) lives next to `pptx/presentation.py`.

### 3.2 Modified files

| Path | Change | Rationale |
|---|---|---|
| `src/pptx/__init__.py` | Add three rows to `content_type_to_part_class_map` (CT.OFC_CUSTOM_PROPERTIES, CT.OFC_CUSTOM_XML_PROPERTIES, and CT.XML → CustomXmlPart **only when partname matches `/customXml/item*.xml`**, see §3.6) | Register part subclasses with the factory |
| `src/pptx/presentation.py` | Add `custom_properties` and `custom_xml_parts` properties | User-facing surface |
| `src/pptx/parts/presentation.py` | Add `custom_properties` lazyproperty, `custom_xml_parts` lazyproperty, helper for "find or create" the parts under the right relationship scope | Where the part-graph wiring lives |
| `src/pptx/package.py` | Add `custom_properties` lazyproperty (mirrors `core_properties`) — package-root scope is correct for `/docProps/custom.xml` per OOXML convention | Package-root relating |
| `src/pptx/oxml/__init__.py` | `register_element_cls(...)` calls for the new oxml classes | Standard registration |
| `src/pptx/types.py` | (Optional) `CustomPropertyValue` type alias for the union | Keep public `__init__.py` clean |
| `pyproject.toml` / `HISTORY.rst` | Bump minor version, log change | Release hygiene |

**No changes** to `src/pptx/opc/constants.py`, `src/pptx/opc/serialized.py`,
`src/pptx/opc/package.py`, or `src/pptx/opc/spec.py`. The constants and content
types we need are already there; the writer auto-derives content types per
part.

### 3.3 Content type and relationship plumbing — no new constants

Verified by reading `src/pptx/opc/constants.py:33–34, 170, 220–229`:

```text
CT.OFC_CUSTOM_PROPERTIES        = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
CT.OFC_CUSTOM_XML_PROPERTIES    = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
CT.XML                          = "application/xml"
RT.CUSTOM_PROPERTIES            = ".../custom-properties"
RT.CUSTOM_XML                   = ".../customXml"
RT.CUSTOM_XML_PROPS             = ".../customXmlProps"
```

`_ContentTypesItem._defaults_and_overrides` (`opc/serialized.py:280–296`) reads
each part's `.content_type` and emits Default-or-Override entries automatically.
Since the `xml` extension already maps to `application/xml` in the default dict,
`/customXml/itemN.xml` (content_type `application/xml`) needs no Override
(Office writes the same way). `/customXml/itemPropsN.xml` becomes an Override
because its content type differs from `application/xml`. `/docProps/custom.xml`
becomes an Override (custom-properties+xml).

### 3.4 Custom-name storage decision

The OOXML spec does **not** define a "name" attribute on a customXml part.
docx-oss's `add_item` stores tags as XML elements; lookup by name there is by
the *element tag*, not the part. We need part-level naming for
`custom_xml_parts.by_name("provenance")`.

**Two options:**

- **(Chosen, default plan)** Store names as a custom document property keyed by
  the part's `datastoreItem_id`: `_pptx_customxml_name_{guid}` → `name`.
  Lossless, round-trips through PowerPoint, no schema invention. Cost: every
  `add(name=...)` also touches `/docProps/custom.xml`.
- (Rejected) Add a `<ds:itemDescription>` child to `itemProps` with a custom
  attribute. Office tolerates it but other tools may strip it; not portable.

**Open question Q3 in §8** — confirm the chosen approach before coding.

### 3.5 Relationship topology — default and override

| Part | Default scope | Source rels file | Override flag | Why |
|---|---|---|---|---|
| `CustomPropertiesPart` (`/docProps/custom.xml`) | package-root | `/_rels/.rels` | none | Office writes it here; sibling of `core.xml` |
| `CustomXmlPart` (`/customXml/itemN.xml`) | presentation-scoped | `/ppt/_rels/presentation.xml.rels` | `scope="package"` on `add(...)` | Office.js / PowerPoint UI only enumerate presentation-scoped customXml |
| `CustomXmlPropertiesPart` (`/customXml/itemPropsN.xml`) | from its CustomXmlPart | `/customXml/_rels/itemN.xml.rels` (always) | n/a | Always a child of the data part |

The `scope` parameter on `CustomXmlParts.add` is the override hatch. Round-trip
test fixtures in §5.3 cover both topologies.

### 3.6 PartFactory ambiguity around `application/xml`

`PartFactory` keys on `content_type` alone. But `/customXml/itemN.xml` has
content_type `application/xml`, which is also the catch-all for unrelated XML
parts. Two options:

- **(Chosen)** Register `CT.XML → Part` (the base class — the existing default
  behavior) and **resolve `CustomXmlPart` by partname pattern** at the
  `PresentationPart.custom_xml_parts` level: enumerate `RT.CUSTOM_XML`
  relationships, wrap each `target_part` in a `CustomXmlPart` facade *if not
  already*. The facade carries the `_element` reference and writes through. No
  new factory ambiguity.
- (Rejected) Subclass `Part` and re-resolve at load time. Risks promoting
  unrelated `application/xml` parts to `CustomXmlPart` and breaking unrelated
  XML-typed parts in third-party PPTX files.

Note: `CustomXmlPropertiesPart` has its own dedicated content type
(`OFC_CUSTOM_XML_PROPERTIES`) so it registers normally with the factory; only
the data part is ambiguous.

### 3.7 Loading existing customXml parts

`OpcPackage._load` already walks every `.rels` file and constructs a `Part` for
every targeted partname (`opc/package.py:240–278`). With `CT.OFC_CUSTOM_XML_PROPERTIES`
mapped to `CustomXmlPropertiesPart` and `CT.OFC_CUSTOM_PROPERTIES` mapped to
`CustomPropertiesPart` in `__init__.py`, those load automatically. The data
part loads as a base `Part` (or `XmlPart` if we pre-register `CT.XML` to
`XmlPart` — TBD; see Q4). The `CustomXmlParts` collection finds them by walking
`RT.CUSTOM_XML` relationships from both the package and the presentation part.

This means files saved by SharePoint, Office.js, or VSTO will round-trip
without code changes — we only need to *enumerate* both relationship sources,
which §3.5 already accounts for.

---

## 4. Compatibility & migration

### 4.1 Backward compatibility

- **No public API removed or renamed.** Two new properties on `Presentation`,
  one new property on `Package`, and four new module files.
- **No change to `[Content_Types].xml` for files that do not use the new
  features.** A presentation produced by code that never touches
  `prs.custom_properties` or `prs.custom_xml_parts` writes byte-equivalent
  output (modulo any unrelated changes).
- **PPTX files containing customXml parts written by other tools** load today
  thanks to mainline's v0.4.1 hotfix; this PR just makes them visible. No
  loader regressions expected — verified by the fixture matrix in §5.3.

### 4.2 Consistency with existing property APIs

`custom_properties` is intentionally shaped to mirror `core_properties`:

| Aspect | `core_properties` | `custom_properties` |
|---|---|---|
| Lazy-create on first access | yes | yes |
| Surface on | `Presentation` and `Package` | `Presentation` and `Package` |
| Underlying part subclass | `CorePropertiesPart` | `CustomPropertiesPart` |
| Per-element setters/getters | typed properties | dict-like, type-dispatched |

The dict-like shape diverges because custom properties are user-keyed, not a
fixed Dublin Core vocabulary. This is the same divergence docx-oss made.

### 4.3 Slide / shape scope — explicitly deferred

PresentationML has a third mechanism: per-slide and per-shape custom data via
`<p:custDataLst><p:tags r:id="rId..."/></p:custDataLst>`, where each `tag`
relationship targets a `tags+xml` part (`CT.PML_TAGS`). This is the mechanism
issue [#578](https://github.com/scanny/python-pptx/issues/578) asks for, and
what `singerla/pptx-automizer` exposes.

This PR does **not** add per-slide or per-shape tag APIs. Reasons:

1. The first consumer (markdown round-trip) wants presentation-scoped data, not
   per-slide.
2. Per-shape `custDataLst` is plumbed differently (slide-rels, not
   presentation-rels) and deserves its own PR with its own API surface.
3. Keeping this PR small reduces review surface and lets the
   presentation-scoped code stabilize first.

A follow-up PR (referenced in §6) will add `Slide.custom_xml_parts` /
`Shape.custom_xml_parts` once this lands.

---

## 5. Testing strategy

### 5.1 Unit tests — oxml classes

**Pattern:** copy `tests/oxml/test_*.py` style. Pure XML in / XML out, no I/O.

| File | Coverage |
|---|---|
| `tests/oxml/test_custom_properties.py` | `CT_CustomProperties` (root), `CT_Property` (each `<property>`), value-type elements (`vt:lpwstr`, `vt:i4`, `vt:bool`, `vt:filetime`, `vt:r8`). Round-trip parse → mutate → serialize. |
| `tests/oxml/test_custom_xml.py` | `CT_DatastoreItem` and `CT_DatastoreSchemaRef`. Verify `itemID` GUID format, schema-ref add/remove. |

### 5.2 Unit tests — parts and collection wrappers

| File | Coverage |
|---|---|
| `tests/parts/test_custom_properties.py` | `CustomPropertiesPart.default(...)`, getter/setter type dispatch, deletion, pid auto-assignment. Uses synthetic XML fixtures the way `tests/parts/test_coreprops.py:1–198` does. |
| `tests/parts/test_custom_xml.py` | `CustomXmlPart.replace_xml(...)`, `datastoreItem_id` round-trip, `schema_refs` add/remove, paired itemProps part lifecycle (add data → itemProps auto-created; remove data → itemProps removed). |
| `tests/test_custom_properties.py` | `CustomProperties` Mapping protocol — `__getitem__`, `__setitem__`, `__delitem__`, `__contains__`, `__iter__`, `__len__`, type dispatch, `set_string`/etc., raises on unknown type. |
| `tests/test_custom_xml.py` | `CustomXmlParts` — `add(...)`, `remove(...)`, `by_name`, `by_guid`, indexing, scope=package vs presentation, name-storage in custom_properties. |

### 5.3 Integration tests — full-package round-trip

Place fixtures under `tests/test_files/customxml/` (new directory). Each fixture
is a real `.pptx` from a different ecosystem; round-trip = open + save + open +
diff payload.

| Fixture | Origin | What it proves |
|---|---|---|
| `sharepoint-saved.pptx` | A presentation saved through SharePoint with VSTO-injected customXml at *package* scope | Loader handles package-root `RT.CUSTOM_XML` and we round-trip without dropping it |
| `officejs-added.pptx` | Office.js `addCustomXmlPart` output (presentation scope) | The "happy path" — Office.js semantics |
| `vsto-document-toolkit.pptx` | A VSTO-tooled deck with `ds:itemID` schema refs | `schema_refs` survive |
| `manual-multipart.pptx` | Hand-crafted with two customXml items + custom properties | N>1 handling |
| `our-output.pptx` | Generated by the test itself using the new API | Sanity check |

Tests:

1. `test_round_trip_preserves_payload` — open, save, re-open; assert
   `custom_xml_parts[i].blob == original_blob` byte-for-byte (modulo lxml
   re-serialization normalization, which is deterministic).
2. `test_round_trip_preserves_topology` — assert package-scope fixtures still
   relate from package root after save; presentation-scope fixtures still relate
   from presentation rels after save.
3. `test_load_with_no_customxml_unchanged` — open a PPTX with no customXml,
   touch nothing, save; assert byte-equivalent (or content-types/rels are at
   least set-equal — see Q5).
4. `test_core_properties_unaffected` — open a PPTX, set both
   `core_properties.author` and `custom_properties["foo"]`; save; re-open;
   assert both round-trip.

### 5.4 Manual verification — PowerPoint UI

The integration test plan does not — and cannot — verify that PowerPoint itself
considers the output legal. Add a manual checklist to the PR description:

- [ ] Open `our-output.pptx` in PowerPoint 365 (Mac and Windows). No repair
      prompt.
- [ ] *File → Properties → Advanced* shows the custom properties.
- [ ] Open in LibreOffice. Document the behavior (LibreOffice preserves
      package-root customXml but historically strips presentation-scoped data
      parts).
- [ ] Open in OnlyOffice / DocumentServer. Document. (See
      [ONLYOFFICE/DocumentServer#1564](https://github.com/ONLYOFFICE/DocumentServer/issues/1564)
      for known gaps.)

The doc page (§6.2) records what we observed so users have realistic
expectations.

### 5.5 Coverage target

Match the project's existing standard (≥95% line coverage on new modules per
`pyproject.toml`). Run with `tox -e py311` (existing tox config).

---

## 6. Documentation

### 6.1 User guide page — `docs/user/custom-xml.rst`

Style: match the other `docs/user/*.rst` pages (`presentations.rst`,
`notes.rst`). Sections:

1. Overview — when to use custom doc properties vs. customXml data parts.
2. Reading and writing custom document properties (with the full type table).
3. Reading and writing customXml data parts (with the string-blob example *and*
   the arbitrary-XML example).
4. Round-trip caveats — what PowerPoint preserves, what LibreOffice / OnlyOffice
   may strip. (Reference §5.4 manual matrix.)
5. Choosing the relationship scope (default vs. `scope="package"` and why).

### 6.2 Dev analysis page — `docs/dev/analysis/customxml.rst`

Match the existing `docs/dev/analysis/*` style (one analysis per OOXML feature,
ECMA-376 references, sample XML, schema diagrams in ASCII). Sections:

1. ECMA-376 references — Part 1 §15.2.4 (Custom XML Data Storage Part) and
   §15.2.12 (Custom File Properties Part).
2. Sample XML for `/docProps/custom.xml`, `/customXml/item1.xml`,
   `/customXml/itemProps1.xml`.
3. Relationship topology diagram — package vs. presentation scope.
4. Why the well-known FMTID is fixed.
5. The `application/xml` content-type ambiguity and how `python-pptx-extended`
   resolves it (§3.6).
6. The `_pptx_customxml_name_<guid>` storage convention (§3.4).

### 6.3 API reference

Add `docs/api/custom_properties.rst` and `docs/api/custom_xml.rst` with the
auto-doc directives. Update `docs/api/presentation.rst` to mention the two new
properties.

### 6.4 HISTORY.rst

A 1.2.0 entry summarizing the feature (the fork's version-bump pattern from
`Plans/review-the-guide-at-swift-kahn.md`).

---

## 7. Phased implementation order

The phases below assume a dedicated `feature/customxml` branch (matches the
`feature/*` branch convention in `git log`).

### Phase 1 — oxml foundation (no public API)

- `src/pptx/oxml/custom_properties.py` — `CT_CustomProperties`, `CT_Property`,
  value-type elements.
- `src/pptx/oxml/custom_xml.py` — `CT_DatastoreItem`, `CT_DatastoreSchemaRef`.
- `src/pptx/oxml/__init__.py` — register element classes.
- Tests: `tests/oxml/test_custom_properties.py`, `tests/oxml/test_custom_xml.py`.

**Deliverable:** new oxml classes parse and serialize round-trip. No
behavior change for callers.

### Phase 2 — Part subclasses

- `src/pptx/parts/custom_properties.py` — `CustomPropertiesPart` with the
  `default` factory and per-element accessors.
- `src/pptx/parts/custom_xml.py` — `CustomXmlPart`, `CustomXmlPropertiesPart`,
  with paired-creation logic (`new_pair(package, ...)`).
- `src/pptx/__init__.py` — register the new content types in
  `content_type_to_part_class_map`.
- Tests: `tests/parts/test_custom_properties.py`, `tests/parts/test_custom_xml.py`.

**Deliverable:** parts load and save correctly when present in a PPTX file.
Still no Presentation-level surface.

### Phase 3 — Public collections and Presentation hooks

- `src/pptx/custom_properties.py` — `CustomProperties` mapping wrapper.
- `src/pptx/custom_xml.py` — `CustomXmlParts` sequence wrapper.
- `src/pptx/parts/presentation.py` — lazyproperties to expose them.
- `src/pptx/package.py` — `custom_properties` lazyproperty.
- `src/pptx/presentation.py` — `custom_properties` and `custom_xml_parts`
  properties.
- Tests: `tests/test_custom_properties.py`, `tests/test_custom_xml.py`,
  `tests/test_presentation.py` additions.

**Deliverable:** end-to-end usage works against the synthetic test fixtures.

### Phase 4 — String-blob helper and integration tests

- `add_string_blob` / `read_string_blob` on `CustomXmlParts`.
- `tests/test_files/customxml/` fixture set (§5.3).
- `tests/integration/test_customxml_roundtrip.py` — end-to-end open-save-reopen.

**Deliverable:** the immediate use case (markdown blob round-trip) is
exercisable from a CLI.

### Phase 5 — Documentation and release

- `docs/user/custom-xml.rst`, `docs/dev/analysis/customxml.rst`, API ref pages.
- `HISTORY.rst` entry, `pyproject.toml` version bump (e.g. `1.1.0` → `1.2.0`).
- Manual PowerPoint UI matrix (§5.4) executed and recorded in the PR
  description.
- Tag and publish (matches the trusted-publishing workflow on the current
  branch).

**Deliverable:** PR ready for principal review.

### Critical-path dependencies

```
Phase 1 ──▶ Phase 2 ──▶ Phase 3 ──▶ Phase 4 ──▶ Phase 5
            (Phase 2 depends on Phase 1's element classes)
            (Phase 3 depends on Phase 2's parts)
            (Phase 4 depends on Phase 3's public API)
```

Phases 1–3 are testable in isolation. Phase 4's fixtures need real third-party
PPTX files, which is why integration tests come last (and partially gate manual
verification, §5.4).

---

## 8. Open questions and decisions for the principal

Numbered for easy reference. Defaults shown so that, if the principal is
indifferent, the plan is unblocked.

**Q1. `add_item(tag, text, **attrs)` shim?**
docx-oss exposes this convenience on `CustomXmlPart`. Useful for "flat list of
items" callers; redundant for callers using `.element` directly.
*Default:* **include it** — low cost, parity with docx-oss, keeps the
"learn-once" surface across the python-openxml family.

**Q2. Distribution name and version.**
Per `Plans/review-the-guide-at-swift-kahn.md`, the fork ships as
`python-pptx-extended` on PyPI. This feature warrants a minor bump (`1.1.0` →
`1.2.0`).
*Default:* **`python-pptx-extended==1.2.0`.** Confirm if you'd rather batch this
with other unreleased fork features instead of a dedicated release.

**Q3. Custom-name storage mechanism (§3.4).**
*Default:* **store names as a reserved custom document property keyed by
datastoreItem GUID** (`_pptx_customxml_name_<guid>`). Lossless, round-trips
through PowerPoint.
*Alternative:* skip name-based lookup entirely; require the caller to track
GUIDs themselves. This is what docx-oss does (no `by_name` on
`custom_xml_parts`). Smaller API, but the markdown-round-trip CLI use case
clearly wants a name.

**Q4. Should `CustomXmlPart` register against `CT.XML`?**
*Default:* **No** — leave `CT.XML` mapping to base `Part` and let the
`CustomXmlParts` collection wrap-on-enumerate (§3.6). Avoids accidentally
upgrading unrelated `application/xml` parts.
*Alternative:* register it and accept the broader scope. Easier loader code,
risk of false positives.

**Q5. Byte-exact preservation of files we do not modify.**
The integration test plan compares payloads, not byte streams (§5.3 test 3).
lxml re-serialization can change attribute order, whitespace, and XML
declaration form even when content is identical.
*Default:* **assert content equivalence (parsed AST equal), not byte
equivalence.** Match scanny upstream's posture.
*Alternative:* invest in a custom serializer that preserves original byte form
for unmodified parts. Significant scope creep; not recommended for this PR.

**Q6. `Slide.custom_xml_parts` / `Shape.custom_xml_parts`.**
*Default:* **out of scope for this PR** (§4.3). Will be a follow-up that
covers `<p:custDataLst><p:tags>` and the slide-rels-rooted topology — issue
#578.
Confirm you agree with deferring this; if you'd rather have one big PR, the
estimate roughly doubles in size (more parts, slide-rels handling, per-shape
API design questions).

**Q7. License headers / attribution to docx-oss.**
The pattern is original to docx-oss (BSD-licensed). MIT (this fork) is
compatible.
*Default:* **add a one-line attribution at the top of `custom_xml.py` and
`custom_properties.py`** noting the docx-oss inspiration with a URL. No code is
copied verbatim; only the API shape is borrowed.

**Q8. Versioning of the generated XML (e.g. provenance metadata about
*python-pptx-extended* itself).**
Some tools stamp the output with a generator hint (e.g.
`<meta generator="python-pptx-extended/1.2.0"/>`). Easy to add to the
string-blob envelope.
*Default:* **don't add this.** Keep the helper minimal; callers who want
provenance write it themselves.

---

## 9. Summary of scope boundaries

**In scope (this PR):**

- Read/write/create custom document properties (`/docProps/custom.xml`).
- Read/write/create customXml data parts (`/customXml/itemN.xml` +
  `/customXml/itemPropsN.xml`) at presentation scope (default) or package
  scope (override).
- String-blob helper for the immediate use case.
- Round-trip safety with files written by SharePoint, Office.js, and VSTO.
- Documentation matching the project's user-guide and analysis-page styles.

**Out of scope (this PR — explicit):**

- Per-slide custom data (`<p:custDataLst><p:tags>`) — issue #578, follow-up PR.
- Per-shape custom data — same follow-up.
- Office.js-style schema validation (we accept `schema_refs` as opaque URIs
  but do not validate payloads against any schema).
- Content controls / structured document tags (SDT) bound to customXml — that's
  a wordprocessingML feature anyway.
- Byte-perfect preservation of files we do not modify (Q5).
- Auto-encoding for `add_string_blob(encoding="base64")` — caller pre-encodes.
- Cross-filesystem name uniqueness checks — `add(name="x")` does not raise if
  another part already has name `"x"`; the principal manages namespacing.

---

## 10. Acceptance — what the principal sees on PR open

1. Branch `feature/customxml` against `ci/pypi-trusted-publishing` (or
   wherever the principal points).
2. ~30 new test cases, ~95% line coverage on the new modules.
3. ~4 fixtures under `tests/test_files/customxml/` documenting the third-party
   topologies.
4. `docs/user/custom-xml.rst` + `docs/dev/analysis/customxml.rst` rendering on
   ReadTheDocs.
5. PR description with the manual-verification matrix from §5.4 filled in.
6. `HISTORY.rst` entry under a new `1.2.0` heading.
7. Examples in the user guide that exercise both the immediate (markdown blob)
   and the general (arbitrary XML) cases.

---

*Plan author: Athena, on behalf of Matthew Horoszowski. Awaiting principal
approval before Phase 1 begins.*
