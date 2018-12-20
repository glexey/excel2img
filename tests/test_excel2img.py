import os

# Run in the tests directory
mypath = os.path.dirname(__file__)
os.chdir(mypath)

def test_all():
    import excel2img
    fnout = "test.png"
    if os.path.exists(fnout): os.unlink(fnout)
    excel2img.export_img("test.xlsx", "test.png", "Sheet1", None)
    assert os.path.exists(fnout)
