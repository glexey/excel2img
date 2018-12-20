import os

# Run in the tests directory
mypath = os.path.dirname(__file__)
os.chdir(mypath)

def test_all():
    import excel2img
    excel2img.export_img("test.xlsx", "test.png", "Sheet1", None)
