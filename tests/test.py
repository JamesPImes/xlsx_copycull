
import os
import unittest
from pathlib import Path

import src.xlsx_copycull as xlsx_copycull


class FileHandler:
    """Helper class for unittest.TestCase."""
    master = Path(r"test_data\test_data.xlsx")
    temp_dir = Path(r"test_temp")
    sheet_name = 'Sheet1'
    temp_fn = 'test_temp.xlsx'
    temp_fp = temp_dir / temp_fn

    def __init__(self):
        self.temp_wbwrapper = None

    def new_copy(self):
        """
        (Setup method.)
        Get a fresh copy.
        :return: The new WorkbookWrapper.
        """
        self.temp_dir.mkdir(exist_ok=True)
        self.temp_wbwrapper = xlsx_copycull.WorkbookWrapper(
            wb_fp=self.master,
            output_filename=self.temp_fn,
            copy_to_dir=self.temp_dir
        )
        return self.temp_wbwrapper

    def reload_wswrapper(self, **kwargs):
        """
        (Setup method.)
        Without reloading the WorkbookWrapper, get rid of the existing
        WorksheetWrapper (if any). Reload a new WorksheetWrapper.
        :param kwargs: Optional kwargs for `.stage_ws()`
        :return: The new WorksheetWrapper.
        """
        for k in self.temp_wbwrapper.ws_dict.copy().keys():
            self.temp_wbwrapper.ws_dict.pop(k)
        wswr = self.temp_wbwrapper.stage_ws(
            ws_name=self.sheet_name,
            header_row=1,
            **kwargs)
        return wswr

    def clean_up(self):
        """
        (Setup method.)
        Close the WorkbookWrapper. Delete the temporary files and
        directory.
        :return: None
        """
        self.temp_wbwrapper.close_wb(save=False)
        self.temp_wbwrapper = None
        for fn in os.listdir(self.temp_dir):
            fp = self.temp_dir / fn
            fp.unlink()
        self.temp_dir.rmdir()


class UnitTest(unittest.TestCase):

    # Helper object for generating new test files and wiping old ones.
    FH = FileHandler()

    master = FH.master
    temp_dir = FH.temp_dir
    sheet_name = FH.sheet_name
    temp_fn = FH.temp_fn
    temp_fp = FH.temp_fp

    def new_copy(self):
        return self.FH.new_copy()

    def reload_wswrapper(self, **kwargs):
        return self.FH.reload_wswrapper(**kwargs)

    def clean_up(self):
        return self.FH.clean_up()

    def test_new_wbwrapper(self):
        wb = self.new_copy()
        self.assertTrue(self.temp_fp.exists())
        self.assertTrue(wb.is_loaded)
        self.clean_up()

    def test_new_wswrapper(self):
        self.new_copy()
        wswr = self.reload_wswrapper()
        self.clean_up()

    def test_rename(self):
        test_name = 'temp'
        wbwp = self.new_copy()
        # Test with wbwp.rename_ws()
        wswp = wbwp.stage_ws(self.sheet_name)
        wbwp.rename_ws(self.sheet_name, test_name)
        wswp = wbwp[test_name]
        self.assertTrue(wswp.ws.title == test_name)
        self.assertFalse(wswp.ws.title == self.sheet_name)
        self.clean_up()

        # Test at stage_ws().
        self.new_copy()
        wswp = self.reload_wswrapper(rename_ws=test_name)
        self.assertTrue(wswp.ws.title == test_name)
        self.assertFalse(wswp.ws.title == self.sheet_name)
        self.clean_up()

        # Test after init.
        self.new_copy()
        wswp = self.reload_wswrapper()
        wswp.rename_ws(test_name)
        self.assertTrue(wswp.ws.title == test_name)
        self.assertFalse(wswp.ws.title == self.sheet_name)

        # Test subscript new name.
        wswp = wbwp[test_name]
        # Test subscript old name.
        with self.assertRaises(KeyError):
            wbwp[self.sheet_name]
        self.clean_up()

    def test_is_loaded(self):
        # Test while open (assert True).
        wbwp = self.new_copy()
        wswp = self.reload_wswrapper()
        self.assertTrue(wbwp.is_loaded)
        self.assertTrue(wswp.is_loaded)
        # Test while closed (assert False).
        wbwp.close_wb(save=False)
        self.assertFalse(wbwp.is_loaded)
        self.assertFalse(wswp.is_loaded)
        self.clean_up()


if __name__ == '__main__':
    fh = FileHandler()
    fh.clean_up()
    unittest.main()
