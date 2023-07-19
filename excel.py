from os import path as _path
from warnings import warn as _warn

import pandas as _pd
import fuckit as _fuckit
import xlwings as _xlwings
import xlwings.main

import funclite.iolib as _iolib


class ExcelAsDataFrame:
    """ Simple wrapper around xlwings to get a named worksheet, range or table as a dataframe.
    Support "with"

    Also exposes xlWings App and workbook objects.

    Members:
        xlBk: An instance of xlwings.Book, exposing the workbook.

    Methods:
        as_df (pandas.DataFrame): Get table, worksheet or range as a dataframe

    Args:
        workbook (str): The workbook to open
        worksheet (str): The worksheet to get data frame
        table (str): Table (listobject) name
        range_ (str): A range name

    Raises:
        UserWarning: If we cannot get or open an xlwings App instance

    Examples:
        Fully specify the worksheet and a range
        >>> with ExcelAsDataFrame('C:/my.xlsx', worksheet='Sheet1', range='A1:C1') as Sheet:
        >>>     Sheet.df
        cola    colb    colc
        1       2       3
        ...

        Ask for a table, cant remember sheet
        >>> with ExcelAsDataFrame('C:/my.xlsx', table='Table1') as Sheet:
        >>>     Sheet.df
        cola    colb    colc
        1       2       3

        Working range of first sheet, should also work with csv
        >>> with ExcelAsDataFrame('C:/my.xlsx') as Sheet:
        >>>     Sheet.df
        cola    colb    colc
        1       2       3

        Just open a workbook, create a new sheet instance to Sheets.mysheet, then create a table from an expanded range on mysheet
        >>> with ExcelAsDataFrame('C:/my.xlsx') as Sheets:
        >>>     Sheets.mysheet = Sheets.xlBk.sheets['Sheet1']
        >>>     Sheets.mysheet.tables.add(source=Sheets.mysheet['A1'].expand(), name='MyTable')

    """
    def __init__(self, workbook: str, worksheet: str = '', table: str = '', range_: str = '', visible: bool = False):
        self._workbook = _path.normpath(workbook)
        self._workbook_file_only = _iolib.get_file_parts2(workbook)[1]
        self._worksheet = worksheet
        self._table = table
        self._range = range_

        try:
            self.xlApp: _xlwings.main.App = _xlwings.apps.active
            if not self.xlApp:
                self.xlApp: _xlwings.main.App = _xlwings.App(visible=visible)
        except:
            self.xlApp: _xlwings.main.App = _xlwings.App(visible=visible)

        if not self.xlApp:
            raise UserWarning('Failed to get or create an xlwings App instance. Is xlwings installed properly?')

        self.xlBk: _xlwings.main.Book = self.xlApp.books.open(self._workbook)
        self.df = self.as_df(worksheet, table, range_) if worksheet or table or range_ else None


    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def close(self):
        """
        Close open excel resources.
        Also writes out the log file.

        Returns: None

        Notes:
            This class should be used as a context handler.
            But if you choose not to, then call instance.close()
            when done.

            This tries pretty much everything to try and close
            excel down.

            All errors are buried under a fuckit.
        """
        with _fuckit:
            bk: _xlwings.main.Book
            for bk in self.xlApp.books:
                bk.close()
                del bk
            self.xlApp.quit()
            self.xlApp.kill()
            xl = _xlwings.apps.active.api
            xl.Quit()

    def as_df(self, worksheet: str = '', table: str = '', range_: str = '', set_as_self_dot_df: bool = True, engine: str = 'xlwings', **kwargs) -> _pd.DataFrame:
        """
        Get an excel sheet used range, listObject (table) or specified range as a pandas dataframe.
        If no table is provided, then gets the worksheet.

        Args:
            worksheet (str): Name of worksheet. If not provided then the first worksheet is assumed when using range
            table (str): Name of the table
            range_ (str): A range
            set_as_self_dot_df (bool): Set self.df to the returned dataframe
            engine (str): "xlwings" or "pandas"
            kwargs: Passed to pandas.read_excel if engine == "pandas"

        Raises:
            ValueError: If both a table and range are provided

        Returns:
            pandas.DataFrame: the table or worksheet as a data frame

        Notes:
            If using the pandas engine, then the args usecols, nrows, and skiprows are generated automatically and do not need to be passed.
        """
        if not engine: engine = 'xlwings'
        if engine not in ['xlwings', 'pandas']:
            _warn('Invalid engine %s specified. Defaulting to "xlwings"' % engine)

        if table and range_:
            raise ValueError('Specify either a table or a range, not both.')

        if table:
            if worksheet:
                r = self.xlApp.books[self._workbook_file_only].sheets[worksheet].tables[table].range
                w = self.xlApp.books[self._workbook_file_only].sheets[worksheet]
            else:
                t, w = self._get_table(table)
                r = t.range
        elif range_:
            if worksheet:
                r = self.xlApp.books[self._workbook_file_only].sheets[worksheet].range(range_)
                w = self.xlApp.books[self._workbook_file_only].sheets[worksheet]
            else:
                r = self.xlApp.books[self._workbook_file_only].sheets[0].range(range_)
                w = self.xlApp.books[self._workbook_file_only].sheets[0]
        else:
            if worksheet:
                r = self.xlApp.books[self._workbook_file_only].sheets[worksheet].used_range
                w = self.xlApp.books[self._workbook_file_only].sheets[worksheet]
            else:
                r = self.xlApp.books[self._workbook_file_only].sheets[0].used_range
                w = self.xlApp.books[self._workbook_file_only].sheets[0]

        if engine.lower() == 'xlwings':
            df = r.expand().options(_pd.DataFrame).value
            df.reset_index(inplace=True)
        else:
            # TODO Debug as_df, particularly this
            df = _pd.read_excel(self._workbook, sheet_name=w.name, nrows=r.last_cell.row - r.row, usecols=tuple(range(r.column-1, r.last_cell.column)), skiprows=r.row-1, **kwargs)

        if set_as_self_dot_df:
            self.df = df
        return df

    def _get_table(self, table: str) -> tuple[xlwings.main.Table, xlwings.main.Sheet]:
        for sheet in self.xlBk.sheets:
            for tbl in sheet.tables:
                if tbl.name.lower() == table.lower():
                    return tbl, sheet

    def table_names(self) -> list[str]:
        """
        Get a list of all tables in the active workbook

        Returns:
            list[str]: List of all tables in the active workbook
        """
        out = []
        for sheet in self.xlBk.sheets:
            for tbl in sheet.tables:
                out += [tbl.name]
        return out

    def view(self) -> None:
        """
        View self.df in using xlwings.view

        Raises:
            UserWarning: If self.df evaluates to False

        Returns: None
        """
        # TODO: debug - gonna have to force the app to visible, otherwise wont work
        if self.df:
            self.xlApp.visible = True
            _xlwings.view(self.df)
        else:
            raise UserWarning('Dataframe not availabe/loaded in your class instance')



def apps_show():
    """Show all xlwings apps, can then close manually.
    Useful for debugging and when not properly closed down during debugging
    """
    if xlwings.apps:
        for app in list(_xlwings.apps):
            try:
                app.visible = True
            except:
                pass
