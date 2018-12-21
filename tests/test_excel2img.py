import os

# Run in the tests directory
mypath = os.path.dirname(__file__)
os.chdir(mypath)

def run_sheet(sheet_name):
    import excel2img
    fnout = sheet_name + ".png"
    if os.path.exists(fnout): os.unlink(fnout)
    try:
        excel2img.export_img("test.xlsx", fnout, sheet_name)
    except OSError as e:
        if "Failed to start Excel" in str(e) and os.environ.get("PYTEST_SKIP_EXCEL"):
            # Waive Excel functionality on Travis
            return
        raise
    assert os.path.exists(fnout), fnout + " didn't get generated"

def test_cells():
    run_sheet('Sheet1')

def test_single_chart():
    run_sheet('Sheet2')

def test_chart_sheet():
    run_sheet('Chart1')

def test_bad_extension():
    import excel2img
    try:
        excel2img.export_img("test.xlsx", "abc.xyz", "Sheet1", None)
    except ValueError as e:
        if 'Unsupported image format' in str(e): return # success
    assert 0, "ValueError('Unsupported image format .XYZ') should have been thrown"
