
Custom XML and Custom Document Properties
==========================================

PowerPoint .pptx packages support two distinct mechanisms for embedding
application-specific structured data alongside slide content:

* **Custom Document Properties** — name/value pairs visible in PowerPoint's UI
  under *File → Properties → Advanced*. Useful for human-readable metadata
  like a build number, source identifier, or workflow status flag.
* **CustomXml data parts** — arbitrary XML payloads with a caller-defined
  namespace. Hidden from end users but preserved by PowerPoint across saves.
  This is the mechanism Office.js, SharePoint, and VSTO add-ins use to attach
  structured application data (provenance, template parameters, audit trails,
  AI-generation markers, etc.).

|pp| exposes both as live, dict-like and sequence-like surfaces on
|Presentation|.

This is a fork-only feature: it is not currently available in upstream
``python-pptx``. See ``Plans/customxml-implementation-plan.md`` in the
repository for the full design rationale.


Custom Document Properties
--------------------------

The :attr:`Presentation.custom_properties` attribute is a Mapping wrapper
around the package's ``/docProps/custom.xml`` part. Read and write it like
a Python ``dict``:

.. code-block:: python

   from pptx import Presentation

   prs = Presentation("input.pptx")

   prs.custom_properties["Source"] = "deck-builder-cli@1.4.2"
   prs.custom_properties["BuildNumber"] = 42
   prs.custom_properties["IsDraft"] = True

   import datetime as dt
   prs.custom_properties["GeneratedAt"] = dt.datetime.now(dt.timezone.utc)

   # Read back
   print(prs.custom_properties["Source"])     # 'deck-builder-cli@1.4.2'
   print("Source" in prs.custom_properties)   # True
   print(list(prs.custom_properties))         # ['Source', 'BuildNumber', ...]

   # Delete
   del prs.custom_properties["IsDraft"]

   prs.save("output.pptx")

Type dispatch on assignment is by Python type:

============================  ========================
Python type                   OOXML element
============================  ========================
``str``                       ``<vt:lpwstr>``
``bool``                      ``<vt:bool>``
``int``                       ``<vt:i4>``
``float``                     ``<vt:r8>``
``datetime.datetime``         ``<vt:filetime>`` (UTC)
============================  ========================

The well-known FMTID ``{D5CDD505-2E9C-101B-9397-08002B2CF9AE}`` is used for
every entry, and the ``pid`` attribute is auto-assigned (≥ 2) per Office's
convention. You don't need to think about either.

For cases where Python's type inference does the wrong thing — for example,
you want the string ``"42"`` rather than the integer 42 — use the explicit
typed setters:

.. code-block:: python

   prs.custom_properties.set_string("NumAsString", "42")
   prs.custom_properties.set_int("Count", 42)
   prs.custom_properties.set_float("Score", 3.14)
   prs.custom_properties.set_bool("Flag", True)
   prs.custom_properties.set_datetime("When", dt.datetime(2026, 1, 1))


CustomXml data parts
--------------------

The :attr:`Presentation.custom_xml_parts` attribute is a sequence-like
collection of customXml data parts attached to the package. Each entry is a
:class:`CustomXmlPart` carrying the user's arbitrary XML payload plus a
sibling :class:`CustomXmlPropertiesPart` carrying the part's ``datastoreItem``
GUID and any declared ``schemaRef`` URIs.

Adding a part
~~~~~~~~~~~~~

.. code-block:: python

   from pptx import Presentation

   prs = Presentation("input.pptx")

   prs.custom_xml_parts.add(
       b'''<?xml version="1.0"?>
       <provenance xmlns="urn:my-app:provenance">
         <source>deck-builder-cli</source>
         <built-at>2026-05-05T14:00:00Z</built-at>
       </provenance>''',
       name="provenance",
       schema_refs=["urn:my-app:provenance"],
   )

   prs.save("output.pptx")

The ``xml`` argument can be ``bytes``, ``str``, or an existing lxml
``_Element``. The ``name`` is an application-assigned label stored in the
custom document properties (under a reserved ``_pptx_customxml_name_*`` key);
it is what :meth:`by_name` looks up.

Lookup
~~~~~~

.. code-block:: python

   prs.custom_xml_parts[0]                            # by index
   prs.custom_xml_parts["item3.xml"]                  # by partname tail
   prs.custom_xml_parts.by_guid("{1A2B3C...}")        # by datastoreItem GUID
   prs.custom_xml_parts.by_name("provenance")         # by user-assigned name

GUID matching is case-insensitive and tolerates the ``{...}`` braces being
present or absent.

Mutation
~~~~~~~~

Each :class:`CustomXmlPart` exposes the live lxml root via ``.element``;
mutating its children mutates the part:

.. code-block:: python

   part = prs.custom_xml_parts.by_name("provenance")
   source = part.element.find("{urn:my-app:provenance}source")
   source.text = "deck-builder-cli@1.4.3"

   prs.save("output.pptx")

To replace the whole payload:

.. code-block:: python

   part.replace_xml(b'<?xml version="1.0"?><new-root xmlns="urn:other"/>')

For the common "flat list of items" shape, ``add_item(tag, text, **attrs)`` is
a one-liner:

.. code-block:: python

   part.add_item("entry", "value", category="meta")

Removal
~~~~~~~

.. code-block:: python

   prs.custom_xml_parts.remove(prs.custom_xml_parts.by_name("provenance"))
   # or
   prs.custom_xml_parts.remove(0)         # by index
   prs.custom_xml_parts.remove("item1.xml")  # by partname tail

``remove`` is idempotent — removing the same part twice is a silent no-op.

The string-blob convenience helper
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For the common case of "stash this string verbatim, give it back to me on
read," |pp| provides a one-shot helper that wraps the string in a reserved
envelope element:

.. code-block:: python

   prs.custom_xml_parts.add_string_blob(
       "readme",
       "# Hello\n\nThis is markdown content embedded in the .pptx.",
       mime_hint="text/markdown",
   )

   # Read back later
   content = prs.custom_xml_parts.read_string_blob("readme")

For binary content, base64-encode at the call site and pass
``encoding="base64"`` so the encoding round-trips:

.. code-block:: python

   import base64
   payload = base64.b64encode(some_bytes).decode("ascii")
   prs.custom_xml_parts.add_string_blob(
       "binary", payload, encoding="base64", mime_hint="application/zip"
   )

The helper does NOT auto-encode for you — encoding is the caller's
responsibility.


Relationship topology — presentation vs. package scope
------------------------------------------------------

OOXML allows a customXml part's relationship to be rooted in either of two
places:

* **Presentation-scoped** — the rel lives in
  ``ppt/_rels/presentation.xml.rels``. This is what Office.js's
  ``addCustomXmlPart`` writes and what PowerPoint's UI surfaces.
* **Package-scoped** — the rel lives in ``_rels/.rels`` (the package root).
  This is the topology VSTO add-ins and SharePoint workflows historically use.

Office.js's ``customXmlParts`` API only enumerates presentation-scoped parts,
so |pp| defaults to that. To match the VSTO/SharePoint topology, pass
``scope="package"``:

.. code-block:: python

   prs.custom_xml_parts.add(b"<vsto/>", name="vsto", scope="package")

The two scopes are not freely interchangeable — once a part is written at one
scope, |pp| preserves that scope on subsequent saves. You can move a part
between scopes by removing and re-adding it.


Round-trip safety
-----------------

Modern PowerPoint preserves customXml parts across saves, including parts
your code did not author. Some other applications behave differently:

* **PowerPoint 365 (Mac and Windows)**: preserves both presentation-scoped
  and package-scoped customXml across edit/save.
* **LibreOffice**: historically preserves package-scoped customXml; behavior
  with presentation-scoped parts is less consistent.
* **OnlyOffice / DocumentServer**: some versions strip customXml on save —
  see `OnlyOffice/DocumentServer issue #1564
  <https://github.com/ONLYOFFICE/DocumentServer/issues/1564>`_.

If your workflow must survive a round-trip through one of these tools, test
with the actual tool before relying on it.

|pp| itself preserves any customXml parts it loads, including those it did
not author — files saved by SharePoint, Office.js, or VSTO load and save
without losing their customXml content.


Choosing between custom properties and customXml parts
------------------------------------------------------

* **Use custom document properties** for small, named, human-readable values
  the user might inspect in PowerPoint's UI.
* **Use customXml parts** for structured data, larger payloads, schema-bound
  XML, or anything you don't want surfaced to end users.

The two mechanisms can coexist — a single .pptx can use both.
