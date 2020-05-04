# -*- coding: utf-8 -*-
#  Copyright 2016 Alexey Gaydyukov
#
#  Licensed under the Apache License, Version 2.0 (the "License");
#  you may not use this file except in compliance with the License.
#  You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
#  Unless required by applicable law or agreed to in writing, software
#  distributed under the License is distributed on an "AS IS" BASIS,
#  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#  See the License for the specific language governing permissions and
#  limitations under the License.

import os
import sys
import win32com.client
from pythoncom import CoInitialize, CoUninitialize
from optparse import OptionParser
from pywintypes import com_error
from PIL import ImageGrab # Note: PIL >= 3.3.1 required to work well with Excel screenshots

class ExcelFile(object):
    @classmethod
    def open(cls, filename):
        obj = cls()
        obj._open(filename)
        return obj

    def __init__(self):
        self.app = None
        self.workbook = None

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()
        return False

    def _open(self, filename):
        excel_pathname = os.path.abspath(filename)  # excel requires abspath
        if not os.path.exists(excel_pathname):
            raise IOError('No such excel file: %s', filename)

        CoInitialize()
        try:
            # Using DispatchEx to start new Excel instance, to not interfere with
            # one already possibly running on the desktop
            self.app = win32com.client.DispatchEx('Excel.Application')
            self.app.Visible = 0
        except:
            raise OSError('Failed to start Excel')

        try:
            self.workbook = self.app.Workbooks.Open(excel_pathname, ReadOnly=True)
            self.workbook.UpdateLink(Name=self.workbook.LinkSources())
        except:
            self.close()
            raise IOError('Failed to open %s'%filename)

    def close(self):
        if self.workbook is not None:
            self.workbook.Close(SaveChanges=False)
            self.workbook = None
        if self.app is not None:
            self.app.Visible = 0
            self.app.Quit()
            self.app = None
        CoUninitialize()


def export_img(fn_excel, fn_image, page=None, _range=None, autofit=None):
    """ Exports images from excel file """

    output_ext = os.path.splitext(fn_image)[1].upper()
    if output_ext not in ('.GIF', '.BMP', '.PNG'):
        raise ValueError('Unsupported image format: %s'%output_ext)

    # if both page and page-less range are specified, concatenate them into range
    if _range is not None and page is not None and '!' not in _range:
        _range = "'%s'!%s"%(page, _range)

    with ExcelFile.open(fn_excel) as excel:
        if _range is None:
            if page is None: page = 1
            try:
                rng = excel.workbook.Sheets(page).UsedRange
                
            except com_error:
                raise Exception("Failed locating used cell range on page %s"%page)
            except AttributeError:
                # This might be a "chart page", try exporting it as a whole
                rng = excel.workbook.Sheets(page).Export(os.path.abspath(fn_image))
                return
            if str(rng) == "None":
                # No used cells on a page. maybe there's a single object.. try simply exporting as png
                shapes = excel.workbook.Sheets(page).Shapes
                if len(shapes) == 1:
                    rng = shapes[0]
                else:
                    raise Exception("Failed locating used cells or single object to print on page %s"%page)
        else:
            try:
                rng = excel.workbook.Application.Range(_range)
            except com_error:
                raise Exception("Failed locating range %s"%(_range))

        if autofit is not None:
            if autofit == True:
                # excel.workbook.Sheets(page).Columns.AutoFit()
                rng.Columns.AutoFit()
            else:
                # excel.workbook.Sheets(page).Columns.AutoFit()
                excel.workbook.Application.Range(autofit).Columns.AutoFit()

        # excel.workbook.Activate() # Trying to solve intermittent CopyPicture failure (didn't work, only becomes worse)
        # rng.Parent.Activate()     # http://answers.microsoft.com/en-us/msoffice/forum/msoffice_excel-msoffice_custom/
        # rng.Select()              # cannot-use-the-rangecopypicture-method-to-copy-the/8bb3ef11-51c0-4fb1-9a8b-0d062bde582b?auth=1
        
        # See http://stackoverflow.com/a/42465354/1924207
        for shape in rng.parent.Shapes: pass

        xlScreen, xlPrinter = 1, 2
        xlPicture, xlBitmap = -4147, 2
        retries, success = 100, False
        while not success:
            try:
                rng.CopyPicture(xlScreen, xlBitmap)
                im = ImageGrab.grabclipboard()
                im.save(fn_image, fn_image[-3:])
                success = True
            except (com_error, AttributeError) as e:
                # http://stackoverflow.com/questions/24740062/copypicture-method-of-range-class-failed-sometimes
                # When other (big) Excel documents are open CopyPicture fails intermittently
                retries -= 1
                #print "CopyPicture failed, retries left:", retries
                if retries == 0: raise


if __name__ == '__main__':

    parser = OptionParser(usage='''%prog excel_filename image_filename [options]\nExamples:
            %prog test.xlsx test.png
            %prog test.xlsx test.png -p Sheet2
            %prog test.xlsx test.png -r MyNamedRange
            %prog test.xlsx test.png -r 'Sheet3!B5:C8'
            %prog test.xlsx test.png -r 'Sheet4!SheetScopedNamedRange'
            ''')
    parser.add_option('-p', '--page', help='pick a page (sheet) by page name. When not specified (and RANGE either not specified or doesn\'t imply a page), first page will be selected')
    parser.add_option('-r', '--range', metavar='RANGE', dest='_range', help='pick a range, in Excel notation. When not specified all used cells on a page will be selected')
    parser.add_option('-F', '--autofit', action='store_true', help='Autofit the column for all range')
    parser.add_option('-f', '--autofit-range', metavar='RANGE', dest='_af_range', help='Autofit the column based on range provided')
    opts, args = parser.parse_args()

    if len(args) != 2:
        parser.print_help(sys.stderr)
        parser.exit()
    
    args_autofit = None
    if opts.autofit is not None:
        args_autofit = True
    elif opts._af_range is not None:
        args_autofit = opts._af_range
    export_img(args[0], args[1], opts.page, opts._range, args_autofit)
