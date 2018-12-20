import os

# Run in the tests directory
mypath = os.path.dirname(__file__)
os.chdir(mypath)

def test_all():
    import excel2img
    fnout = "test.png"
    if os.path.exists(fnout): os.unlink(fnout)
    try:
        excel2img.export_img("test.xlsx", "test.png", "Sheet1", None)
    except OSError as e:
        if "Failed to start Excel" in str(e) and os.environ.get("PYTEST_SKIP_EXCEL"):
            # Waive Excel functionality on Travis
            return
    assert os.path.exists(fnout)
