excel2img
=========

Save ranges from Excel documents as images

Requirements
------------

1. Python 2.7, 3.3 and later
2. `pywin32 <http://sourceforge.net/projects/pywin32/files/pywin32>`__
3. `Pillow <https://pypi.python.org/pypi/Pillow>`__ >= 3.3.1
4. Microsoft Excel (tested with Office 2013, on Windows 10)

Usage as python module
----------------------

.. code:: python

    import excel2img

    # Save as PNG the range of used cells in test.xlsx on page named "Sheet1"
    excel2img.export_img("test.xlsx", "test.png", "Sheet1", None)

    # Save as BMP the range B2:C15 in test.xlsx on page named "Sheet2"
    excel2img.export_img("test.xlsx", "test.bmp", "", "Sheet2!B2:C15")

    # Save as GIF the range "MyNamedRange"
    excel2img.export_img("test.xlsx", "test.gif", "", "MyNamedRange")

Usage from command line
-----------------------

.. code:: shell

    # Save as PNG the range of used cells in test.xlsx on first page
    python excel2img.py test.xlsx test.png

    # Save as PNG the range of used cells in test.xlsx on page "Sheet2"
    python excel2img.py test.xlsx test.png -p Sheet2

    # Save as PNG the range "MyNamedRange"
    python excel2img.py test.xlsx test.png -r MyNamedRange

    # More range syntax examples
    python excel2img.py test.xlsx test.gif -r 'Sheet3!B5:C8'
    python excel2img.py test.xlsx test.bmp -r 'Sheet4!SheetScopedNamedRange'

Author
=======

Alexey Gaydyukov <glexey@gmail.com>

License
========
Apache License 2.0

Credits
========
Inspired by `visio2img <https://github.com/visio2img/visio2img>`__

