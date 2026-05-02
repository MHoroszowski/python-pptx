Fork notice
-----------

This distribution, ``python-pptx-extended``, is a fork of
`scanny/python-pptx`_ at upstream version 1.0.2. The import name is unchanged
(``import pptx``), so existing user code continues to work. The fork adds the
following features on top of upstream:

- Full shadow effect API on ``ShadowFormat`` (outer/inner/preset shadows).
- Bullet and numbered list formatting on paragraphs.
- Per-edge border styling for table cells.
- ``cap_style`` and ``join_style`` properties on ``LineFormat``.
- Line-end shape types (arrow / triangle / oval / etc.).

Because the import package name (``pptx``) is shared with the upstream
distribution, ``python-pptx`` and ``python-pptx-extended`` cannot be installed
into the same environment — install one or the other.

.. _`scanny/python-pptx`: https://github.com/scanny/python-pptx

About python-pptx
-----------------

*python-pptx* is a Python library for creating, reading, and updating PowerPoint (.pptx)
files.

A typical use would be generating a PowerPoint presentation from dynamic content such as
a database query, analytics output, or a JSON payload, perhaps in response to an HTTP
request and downloading the generated PPTX file in response. It runs on any Python
capable platform, including macOS and Linux, and does not require the PowerPoint
application to be installed or licensed.

It can also be used to analyze PowerPoint files from a corpus, perhaps to extract search
indexing text and images.

In can also be used to simply automate the production of a slide or two that would be
tedious to get right by hand, which is how this all got started.

More information is available in the `python-pptx documentation`_.

Browse `examples with screenshots`_ to get a quick idea what you can do with
python-pptx.

.. _`python-pptx documentation`:
   https://python-pptx.readthedocs.org/en/latest/

.. _`examples with screenshots`:
   https://python-pptx.readthedocs.org/en/latest/user/quickstart.html
