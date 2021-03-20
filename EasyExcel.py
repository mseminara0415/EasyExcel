import win32com.client as win32
import os
import re
import sys
import shutil


class EasyExcel:
    """
    EasyExcel acts as a wrapper for Pywin32, allowing for straight forward Excel editing/formatting.

    """

    @staticmethod
    def initialize_excel(visible: bool, display_alerts: bool, screen_updating: bool, enable_events: bool):
        """
        Initialize Pywin Excel application.
        :param visible:
        :param display_alerts:
        :param screen_updating:
        :param enable_events:
        :return:
        """
        # Set up Excel Workbook
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = visible
            excel.Application.DisplayAlerts = display_alerts
            excel.ScreenUpdating = screen_updating
            excel.EnableEvents = enable_events
        except AttributeError:
            # Remove cache and try again.
            MODULE_LIST = [m.__name__ for m in sys.modules.values()]
            for module in MODULE_LIST:
                if re.match(r'win32com\.gen_py\..+', module):
                    del sys.modules[module]
            shutil.rmtree(
                os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = visible
            excel.Application.DisplayAlerts = display_alerts
            excel.ScreenUpdating = screen_updating
            excel.EnableEvents = enable_events

        return excel

    def __init__(self, file_path: str):
        self.filepath = file_path
        self.excel = self.initialize_excel(False, False, False, False)
        self.wb = self.excel.Workbooks.Open(self.filepath)

    @property
    def sheets(self):
        """
        Get all sheets within workbook.
        :return:
        """
        return [sheet for sheet in self.wb.Sheets]

    def color_scale(self, worksheet: str, cell_range_start=None,
                    cell_range_end=None):
        """
        Add in color scale for selected range.
        :param worksheet:
        :param cell_range_start:
        :param cell_range_end:
        :return:
        """

        sheet = self.wb.Worksheets(worksheet)

        # Create Working Cell Range
        if cell_range_end is not None:
            working_range = sheet.Range(
                f"{cell_range_start}:{cell_range_end}")
        else:
            working_range = sheet.Range(f"{cell_range_start}")

        working_range.FormatConditions.AddColorScale(ColorScaleType=3)

    def merge_cells(self, worksheet, cell_range_start=None, cell_range_end=None,
                    center_text=True):
        """
        Merge Excel Cell Range. Also Provides an option to center align.
        :return:
        """
        sheet = self.wb.Worksheets(worksheet)

        # Create Working Cell Range
        if cell_range_end is not None:
            working_range = sheet.Range(
                f"{cell_range_start}:{cell_range_end}")
        else:
            working_range = sheet.Range(f"{cell_range_start}")

        working_range.MergeCells = True

        # Center Text
        if center_text:
            sheet.Range(f"{cell_range_start}").HorizontalAlignment = -4108

    def bold_cells(self, worksheet, cell_range_start=None, cell_range_end=None):
        """
        Bold selected cell range.
        :param worksheet:
        :param cell_range_start:
        :param cell_range_end:
        :return:
        """

        sheet = self.wb.Worksheets(worksheet)

        # Create Working Cell Range
        if cell_range_end is not None:
            working_range = sheet.Range(
                f"{cell_range_start}:{cell_range_end}")
        else:
            working_range = sheet.Range(f"{cell_range_start}")

        working_range.Font.Bold = True

    def close_workbook(self):
        """
        Close and Save workbook.
        :return:
        """
        self.wb.Close(True)
        self.wb = None
        self.excel = None
