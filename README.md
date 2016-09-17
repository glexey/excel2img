# excel2img
Save ranges from Excel documents as images

## Usage as python module

```python
import excel2img

# Save as PNG the range of used cells in test.xlsx on page named "Sheet1"
excel2img.export_img("test.xlsx", "test.png", "Sheet1", None)

# Save as BMP the range B2:C15 in test.xlsx on page named "Sheet2"
excel2img.export_img("test.xlsx", "test.bmp", "", "Sheet2!B2:C15")

# Save as GIF the range "MyNamedRange"
excel2img.export_img("test.xlsx", "test.gif", "", "MyNamedRange")
```

## Usage from command line

```shell
# Save as PNG the range of used cells in test.xlsx on first page
python excel2img.py test.xlsx test.png

# Save as PNG the range of used cells in test.xlsx on page "Sheet2"
python excel2img.py test.xlsx test.png -p Sheet2

# Save as PNG the range "MyNamedRange"
python excel2img.py test.xlsx test.png -r MyNamedRange

# More range syntax examples
python excel2img.py test.xlsx test.gif -r 'Sheet3!B5:C8'
python excel2img.py test.xlsx test.bmp -r 'Sheet4!SheetScopedNamedRange'
```

## Requirements
1. Python 2.7.x (TODO: add support for 3.x)
1. Pillow >= 3.3.1 (pip install PIL --upgrade)
1. win32com
1. Microsoft Excel (tested with Office 2013, on Windows 10)
