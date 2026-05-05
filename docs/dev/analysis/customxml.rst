.. _CustomXml:

CustomXml and Custom Document Properties
=========================================

Two distinct OOXML mechanisms support embedding application-specific data in
a ``.pptx`` package:

1. **Custom Document Properties** at ``/docProps/custom.xml`` — visible in
   PowerPoint UI under *File → Properties → Advanced*. ECMA-376 Part 1 §15.2.12.
2. **CustomXml data parts** at ``/customXml/itemN.xml`` paired with
   ``/customXml/itemPropsN.xml`` — hidden from end users; the mechanism
   Office.js, SharePoint workflows, and VSTO add-ins use to embed structured
   data. ECMA-376 Part 1 §15.2.4.


Custom Document Properties
--------------------------

XML specimen
~~~~~~~~~~~~

.. highlight:: xml

::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <Properties
      xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
      xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Source">
      <vt:lpwstr>deck-builder-cli@1.4.2</vt:lpwstr>
    </property>
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="BuildNumber">
      <vt:i4>42</vt:i4>
    </property>
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="IsDraft">
      <vt:bool>true</vt:bool>
    </property>
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="GeneratedAt">
      <vt:filetime>2026-05-05T14:00:00Z</vt:filetime>
    </property>
  </Properties>

Notable details
~~~~~~~~~~~~~~~

* The ``fmtid`` attribute is the same well-known GUID
  ``{D5CDD505-2E9C-101B-9397-08002B2CF9AE}`` for every user-defined property.
  Office uses different FMTIDs for system-defined property sets (e.g. SharePoint
  fields), but |pp| writes the user-defined FMTID exclusively.
* ``pid`` values 0 and 1 are reserved by the spec; user properties start at 2.
  |pp| auto-assigns the next free integer ≥ 2 within the part.
* The typed value child belongs to the ``vt:`` namespace
  (``http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes``).
  Five types are supported: ``lpwstr`` (Unicode string), ``i4`` (32-bit signed
  int), ``r8`` (IEEE-754 double), ``bool``, and ``filetime``
  (ISO-8601 UTC, ``Z``-suffixed).


CustomXml data parts
--------------------

Each customXml entry is a **pair** of parts: one for the user's arbitrary XML
payload and one for the metadata about it.

XML specimen — data part
~~~~~~~~~~~~~~~~~~~~~~~~

The data part at ``/customXml/item1.xml`` carries arbitrary XML the application
chose to embed. The root element name and namespace are caller-defined::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <provenance xmlns="urn:my-app:provenance">
    <source>deck-builder-cli</source>
    <built-at>2026-05-05T14:00:00Z</built-at>
  </provenance>

The content type is ``application/xml`` — the OPC default for the ``xml``
extension, so no per-part Override entry is written into ``[Content_Types].xml``.

XML specimen — itemProps part
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The sibling at ``/customXml/itemProps1.xml`` carries the ``datastoreItem`` GUID
that uniquely identifies the data part across edits, plus an optional
``schemaRefs`` list declaring the namespaces the data part claims to conform
to::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <ds:datastoreItem
      xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"
      ds:itemID="{1A2B3C4D-5E6F-7890-ABCD-EF1234567890}">
    <ds:schemaRefs>
      <ds:schemaRef ds:uri="urn:my-app:provenance"/>
    </ds:schemaRefs>
  </ds:datastoreItem>

Content type ``application/vnd.openxmlformats-officedocument.customXmlProperties+xml``
is written as an Override entry for this partname.

Relationship topology
~~~~~~~~~~~~~~~~~~~~~

The data part's relationship can be rooted in either the package or the
presentation::

   PRESENTATION-SCOPED (default; what Office.js writes)
   ────────────────────────────────────────────────────
   /ppt/_rels/presentation.xml.rels
       └─ Type=customXml ─▶ /customXml/item1.xml
                                └─ /customXml/_rels/item1.xml.rels
                                       └─ Type=customXmlProps ─▶ /customXml/itemProps1.xml


   PACKAGE-SCOPED (VSTO / SharePoint topology)
   ───────────────────────────────────────────
   /_rels/.rels
       └─ Type=customXml ─▶ /customXml/item1.xml
                                └─ /customXml/_rels/item1.xml.rels
                                       └─ Type=customXmlProps ─▶ /customXml/itemProps1.xml

The two scopes are not interchangeable — Office.js's ``customXmlParts``
collection only enumerates presentation-scoped parts (see this
`Microsoft Q&A response
<https://learn.microsoft.com/en-us/answers/questions/5586825/how-to-add-a-proper-customxml-to-a-powerpoint-pres>`_).

|pp| defaults to presentation-scoped to match Office.js. The
``scope="package"`` parameter on
:meth:`pptx.custom_xml.CustomXmlParts.add` is the override hatch for VSTO /
SharePoint compatibility.


Design decisions
----------------

The ``application/xml`` content-type ambiguity
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

``PartFactory.part_type_for`` keys on content type alone, but ``application/xml``
is the catch-all default for the ``xml`` extension — every customXml data part
shares it with potentially-unrelated XML parts in third-party packages.

|pp| chooses to **not** register :class:`CustomXmlPart` against ``application/xml``.
Loaded data parts arrive as base ``Part`` instances; the
:class:`CustomXmlParts` collection upgrades them to :class:`CustomXmlPart`
in-place via ``__class__`` swap on first enumeration. This avoids accidentally
promoting unrelated ``application/xml`` parts in third-party packages.

The custom-name storage convention
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

OOXML does not define a "name" attribute on customXml parts. To support
``custom_xml_parts.by_name("provenance")``, |pp| stores user-assigned names
as reserved entries in the custom document properties part keyed by the
data part's ``datastoreItem`` GUID:

::

  <op:property name="_pptx_customxml_name_{1A2B...}" pid="...">
    <vt:lpwstr>provenance</vt:lpwstr>
  </op:property>

This is lossless, round-trips through PowerPoint, and requires no schema
invention. The reserved entries are visible in PowerPoint's
*File → Properties → Advanced* UI by design — what the user sees in the app
matches what the Python API exposes.

Round-trip safety with third-party tools
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

PowerPoint 365 (Mac and Windows) preserves both topologies across edits.
LibreOffice historically preserves package-scoped parts but is less
consistent with presentation-scoped data parts. OnlyOffice / DocumentServer
strips customXml on save in some versions
(`OnlyOffice issue #1564 <https://github.com/ONLYOFFICE/DocumentServer/issues/1564>`_).

|pp| preserves any customXml part it loads, including those it did not
author — files saved by SharePoint, Office.js, or VSTO add-ins load and save
without losing their customXml content.


References
----------

* `ECMA-376 Part 1, §15.2.4 — Custom XML Data Storage Part <https://ecma-international.org/publications-and-standards/standards/ecma-376/>`_
* `ECMA-376 Part 1, §15.2.12 — Custom File Properties Part <https://ecma-international.org/publications-and-standards/standards/ecma-376/>`_
* `MS Q&A on presentation- vs. package-scoped customXml topology <https://learn.microsoft.com/en-us/answers/questions/5586825/how-to-add-a-proper-customxml-to-a-powerpoint-pres>`_
* `Office.js CustomXmlPart API <https://learn.microsoft.com/en-us/javascript/api/office/office.customxmlpart>`_
* `python-docx-oss custom-xml docs <https://python-docx-oss.readthedocs.io/en/latest/user/custom-xml.html>`_ (the docx-equivalent pattern, which |pp|'s API mirrors)
