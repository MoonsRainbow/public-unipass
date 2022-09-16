import sys
import tkinter
import datetime
import calendar
from tkinter import (ttk)
from tkinter.font import Font
from tkcalendar import DateEntry
from View.Resource.style import APP_FONT
from View.Resource.dictionary import HEAD_CATEGORY


class Navigation(tkinter.Frame):
    def __init__(self, _app: tkinter.Tk):
        self._code = 'Navigation'
        self.APP = _app
        super().__init__(self.APP, bg='#DDDDDD')
        self.place(rely=0.06, relheight=0.94, relwidth=0.06)
        self.BUTTONS = {}

    def insert_button(self, _id: str, _text: str, _command, _side=None):
        _button = tkinter.Button(
            master=self,
            name=_id,
            text=_text,
            relief=tkinter.FLAT,
            overrelief=tkinter.SOLID,
            command=_command,
            repeatdelay=1000,
            repeatinterval=100,
            height=3
        )
        if _side is None:
            _button.pack(fill=tkinter.X)
        else:
            _button.pack(fill=tkinter.X, side=_side)
        self.BUTTONS[_id] = _button

    def insert_divide(self):
        _frame = tkinter.Frame(self, height=1, bd=0)
        _frame.pack(fill=tkinter.X, ipady=10)
        ttk.Separator(_frame, orient=tkinter.HORIZONTAL).place(relwidth=1, rely=0.48)

    def disable_button(self, _id):
        for _key, _button in self.BUTTONS.items():
            if _id == _key:
                _button.configure(state=tkinter.DISABLED)
            else:
                _button.configure(state=tkinter.NORMAL)


class Header(tkinter.Frame):
    def __init__(self, _app: tkinter.Tk):
        self._code = 'Header'
        self.APP = _app
        super().__init__(self.APP, bg='#DDDDDD')
        self.place(relheight=0.06, relwidth=1)
        self.TITLE = tkinter.StringVar(master=self, value='')
        self.MAN_NAME = tkinter.StringVar(master=self, value='')
        tkinter.Label(self, bg='#DDDDDD', text='UNIPASS', anchor=tkinter.CENTER).pack(fill=tkinter.Y, padx=10, side=tkinter.LEFT)
        tkinter.Label(self, bg='#DDDDDD', anchor=tkinter.CENTER, textvariable=self.TITLE).pack(fill=tkinter.Y, padx=30, side=tkinter.LEFT)
        tkinter.Label(self, bg='#DDDDDD', textvariable=self.MAN_NAME).pack(fill=tkinter.Y, padx=10, side=tkinter.RIGHT)

    def set_man_name(self, _name):
        self.MAN_NAME.set(_name + ' 님 안녕하세요.')

    def set_title(self, _title):
        self.TITLE.set(_title)


class Body(tkinter.Frame):
    def __init__(self, _parent):
        self._code = 'Body'
        self.APP = _parent
        super().__init__(self.APP, padx=10, pady=10)

    def body_place(self):
        self.place(relx=0.06, rely=0.06, relwidth=0.94, relheight=0.94)

    def unbind_widget(self):
        self.APP.unbind('<Return>')
        self.APP.unbind('<Configure>')
        self.APP.unbind('<Control-c>')
        self.APP.unbind('<Button-1>')
        self.APP.unbind('<Control-1>')
        self.APP.unbind('<Shift-1>')
        self.APP.unbind('<Double-1>')

    def destroy_widget(self):
        self.unbind_widget()
        for _widget in [_w for _w in self.children.values()]:
            _widget.destroy()


class TitleEntry(tkinter.Frame):
    def __init__(self, _parent, _title, _is_pw: bool = False, _label_size: float = 2):
        self._code = 'TitleEntry'
        super().__init__(_parent, padx=5, pady=5)
        tkinter.Label(self, text=_title, foreground='gray').place(relheight=1, relwidth=float(_label_size / 10))
        self.VALUE = tkinter.StringVar(master=self, value='')
        if _is_pw:
            tkinter.Entry(self, textvariable=self.VALUE, show='●').place(relx=float(_label_size / 10), relheight=1, relwidth=1 - float(_label_size / 10))
        else:
            tkinter.Entry(self, textvariable=self.VALUE).place(relx=float(_label_size / 10), relheight=1, relwidth=1 - float(_label_size / 10))

    def get_value(self):
        return self.VALUE.get()

    def set_value(self, _str: str):
        self.VALUE.set(_str)

    def clear_value(self):
        self.VALUE.set('')


class LabelText(tkinter.LabelFrame):
    def __init__(self, _parent, _title, _var):
        self._code = 'LabelText'
        super().__init__(_parent, text=_title, padx=5, pady=5, foreground='gray')
        self.LABEL = tkinter.Label(self, textvariable=_var)
        self.LABEL.pack(fill=tkinter.Y, expand=tkinter.TRUE)


class LabelEntry(tkinter.LabelFrame):
    def __init__(self, _parent, _title):
        self._code = 'LabelEntry'
        super().__init__(_parent, text=_title, padx=5, pady=5, foreground='gray')
        self.VALUE = tkinter.StringVar(master=self, value='')
        self.ENTRY = tkinter.Entry(self, textvariable=self.VALUE)
        self.ENTRY.pack(fill=tkinter.BOTH, expand=tkinter.TRUE)

    def get_value(self):
        return self.VALUE.get()

    def set_value(self, _str: str):
        self.VALUE.set(_str)

    def clear_value(self):
        self.VALUE.set('')


class LabelButton(tkinter.LabelFrame):
    def __init__(self, _parent, _title, _b_text, _command):
        self._code = 'LabelButton'
        super().__init__(_parent, text=_title, padx=5, pady=5, foreground='gray')
        self.BUTTON = tkinter.Button(self, text=_b_text, overrelief=tkinter.SOLID, command=_command, repeatdelay=1000, repeatinterval=100)
        self.BUTTON.pack(fill=tkinter.BOTH, expand=tkinter.TRUE)


class LabelCombo(tkinter.LabelFrame):
    def __init__(self, _parent, _title, _values: dict):
        self._code = 'LabelCombo'
        super().__init__(_parent, text=_title, padx=5, pady=5, foreground='gray')
        self.pack(side=tkinter.LEFT, fill=tkinter.Y)
        self.VALUES = _values
        self.COMBO = ttk.Combobox(self, values=list(self.VALUES.keys()), state='readonly')
        self.COMBO.place(relheight=1, relwidth=1)
        self.COMBO.current(0)

    def set_combo_selected_command(self, _selected_command = None):
        if _selected_command is not None:
            self.COMBO.bind('<<ComboboxSelected>>', _selected_command)
        else:
            self.COMBO.unbind('<<ComboboxSelected>>')

    def set_combo_state(self, _state: [tkinter.DISABLED, str] = 'readonly'):
        self.COMBO.configure(state=_state)

    def get_value_text(self):
        return list(self.VALUES.keys())[self.COMBO.current()]

    def get_value_index(self):
        return self.VALUES.get(list(self.VALUES.keys())[self.COMBO.current()])

    def clear_value(self):
        self.COMBO.current(0)


class LabelCalendar(tkinter.LabelFrame):
    def __init__(self, _parent, _title, _is_on_off: bool = False):
        self._code = 'LabelCalendar'
        self.START_YEAR = int(datetime.datetime.now().year)
        self.CLOSE_YEAR = int(datetime.datetime.now().year)
        self.START_MONTH = int(datetime.datetime.now().month) - 1
        self.CLOSE_MONTH = int(datetime.datetime.now().month) + 1
        self.START_DAY = 1
        self.CLOSE_DAY = int(calendar.monthrange(self.CLOSE_YEAR, self.CLOSE_MONTH)[1])
        super().__init__(_parent, text=_title, padx=5, pady=5, foreground='gray')
        self.START_CALENDAR = DateEntry(
            self,
            date_pattern='yyyy-MM-dd',
            background='white',
            foreground='#DDDDDD',
            year=self.START_YEAR,
            month=self.START_MONTH,
            day=self.START_DAY
        )
        self.CLOSE_CALENDAR = DateEntry(
            self,
            date_pattern='yyyy-MM-dd',
            background='white',
            foreground='#DDDDDD',
            year=self.CLOSE_YEAR,
            month=self.CLOSE_MONTH,
            day=self.CLOSE_DAY
        )
        self.CHECK_STATUS = tkinter.BooleanVar(master=self, value=False)
        self.USE_FLAG = True
        if _is_on_off:
            ttk.Checkbutton(self, text='', variable=self.CHECK_STATUS, command=self._check_click).place(relwidth=0.1, relheight=1, relx=0)
            self.START_CALENDAR.place(relwidth=0.4, relheight=1, relx=0.1)
            tkinter.Label(self, text='~', foreground='gray').place(relwidth=0.1, relheight=1, relx=0.5)
            self.CLOSE_CALENDAR.place(relwidth=0.4, relheight=1, relx=0.6)
            self._check_click()
        else:
            self.START_CALENDAR.place(relwidth=0.45, relheight=1, relx=0)
            tkinter.Label(self, text='~', foreground='gray').pack(fill=tkinter.Y, expand=tkinter.TRUE)
            self.CLOSE_CALENDAR.place(relwidth=0.45, relheight=1, relx=0.55)

    def _check_click(self):
        if self.CHECK_STATUS.get():
            self.START_CALENDAR.configure(state=tkinter.NORMAL)
            self.CLOSE_CALENDAR.configure(state=tkinter.NORMAL)
            self.USE_FLAG = True
        else:
            self.START_CALENDAR.configure(state=tkinter.DISABLED)
            self.CLOSE_CALENDAR.configure(state=tkinter.DISABLED)
            self.USE_FLAG = False

    @staticmethod
    def validate_date(_data: str):
        try:
            datetime.datetime.strptime(_data, '%Y-%m-%d')
            return True
        except ValueError:
            return False

    def get_start_value(self):
        if self.USE_FLAG:
            if self.validate_date(self.START_CALENDAR.get()):
                return str(self.START_CALENDAR.get())
            else:
                self.START_CALENDAR.set_date(
                    '{}-{}-{}'.format(
                        self.START_YEAR,
                        str(self.START_MONTH).zfill(2),
                        str(self.START_DAY).zfill(2)
                    )
                )
                return str(self.START_CALENDAR.get())
        else:
            return ''

    def get_close_value(self):
        if self.USE_FLAG:
            if self.validate_date(self.CLOSE_CALENDAR.get()):
                return str(self.CLOSE_CALENDAR.get())
            else:
                self.CLOSE_CALENDAR.set_date(
                    '{}-{}-{}'.format(
                        self.CLOSE_YEAR,
                        str(self.CLOSE_MONTH).zfill(2),
                        str(self.CLOSE_DAY).zfill(2)
                    )
                )
                return str(self.CLOSE_CALENDAR.get())
        else:
            return ''

    def clear_value(self):
        if not self.USE_FLAG:
            self.START_CALENDAR.configure(state=tkinter.NORMAL)
            self.CLOSE_CALENDAR.configure(state=tkinter.NORMAL)
        self.START_CALENDAR.set_date(
            '{}-{}-{}'.format(
                self.START_YEAR,
                str(self.START_MONTH).zfill(2),
                str(self.START_DAY).zfill(2)
            )
        )
        self.CLOSE_CALENDAR.set_date(
            '{}-{}-{}'.format(
                self.CLOSE_YEAR,
                str(self.CLOSE_MONTH).zfill(2),
                str(self.CLOSE_DAY).zfill(2)
            )
        )
        if not self.USE_FLAG:
            self.START_CALENDAR.configure(state=tkinter.DISABLED)
            self.CLOSE_CALENDAR.configure(state=tkinter.DISABLED)


class LabelEntryButton(tkinter.LabelFrame):
    def __init__(self, _parent, _title, _b_text, _command, _readonly=False):
        self._code = 'LabelEntryButton'
        super().__init__(_parent, text=_title, padx=5, pady=5, foreground='gray')
        self.VALUE = tkinter.StringVar(master=self, value='')
        if _readonly:
            self.ENTRY = tkinter.Entry(self, textvariable=self.VALUE, state='readonly')
        else:
            self.ENTRY = tkinter.Entry(self, textvariable=self.VALUE)
        self.ENTRY.pack(fill=tkinter.BOTH, expand=tkinter.TRUE, side=tkinter.LEFT)
        self.BUTTON = tkinter.Button(self, text=_b_text, overrelief=tkinter.SOLID, command=_command, repeatdelay=1000, repeatinterval=100, width=10)
        self.BUTTON.pack(side=tkinter.RIGHT, fill=tkinter.Y)

    def get_value(self):
        return self.VALUE.get()

    def set_value(self, _str: str = ''):
        self.VALUE.set(_str)

    def clear_value(self):
        self.VALUE.set('')


class LabelEnvDownloadPath(tkinter.LabelFrame):
    def __init__(self, _parent, _title, _desc, _var, _b_text, _command, _readonly: bool = True):
        self._code = 'LabelEnv'
        super().__init__(_parent, text=_title, padx=5, pady=5, foreground='gray')
        self.DESC_LABEL = tkinter.Label(self, text=_desc, pady=10)
        if _readonly:
            self.APP_VALUE = _var
            self.VALUE = None
            self.VALUE_LABEL = tkinter.Label(self, textvariable=self.APP_VALUE, background='white', anchor=tkinter.W, padx=10)
        else:
            self.APP_VALUE = _var
            self.VALUE = tkinter.StringVar(self, value='')
            self.VALUE_LABEL = tkinter.Label(self, textvariable=self.VALUE)
        self.BUTTON = tkinter.Button(self, text=_b_text, overrelief=tkinter.SOLID, command=_command, repeatdelay=1000, repeatinterval=100, width=15)
        self.DESC_LABEL.pack(side=tkinter.TOP, anchor=tkinter.W)
        self.VALUE_LABEL.pack(side=tkinter.LEFT, fill=tkinter.BOTH, expand=True)
        self.BUTTON.pack(side=tkinter.RIGHT, fill=tkinter.Y)


class Pagination(tkinter.LabelFrame):
    def __init__(self, _parent, _pn, _ep, _prev_command, _move_command, _next_command):
        self._code = 'Pagination'
        super().__init__(_parent, text='페이지 이동', pady=5, foreground='gray')
        self.TABLE_PAGE_NO = _pn
        self.PAGE_NO = tkinter.StringVar(self, '1')
        self.PAGE_NO_ENTRY = tkinter.Entry(self, textvariable=self.PAGE_NO, width=5)
        self.END_PAGE_LABEL = tkinter.Label(self, textvariable=_ep)
        self.PREV_BUTTON = tkinter.Button(self, text='←', overrelief='solid', command=_prev_command, repeatdelay=1000, repeatinterval=100)
        self.MOVE_BUTTON = tkinter.Button(self, text='이동', overrelief='solid', command=_move_command, repeatdelay=1000, repeatinterval=100)
        self.NEXT_BUTTON = tkinter.Button(self, text='→', overrelief='solid', command=_next_command, repeatdelay=1000, repeatinterval=100)
        self.PREV_BUTTON.pack(fill=tkinter.BOTH, side=tkinter.LEFT, padx=5, expand=tkinter.TRUE)
        self.PAGE_NO_ENTRY.pack(fill=tkinter.BOTH, side=tkinter.LEFT, expand=tkinter.TRUE)
        self.MOVE_BUTTON.pack(fill=tkinter.BOTH, side=tkinter.LEFT, expand=tkinter.TRUE)
        tkinter.Label(self, text='/', foreground='gray').pack(fill=tkinter.Y, side=tkinter.LEFT, expand=tkinter.TRUE)
        self.END_PAGE_LABEL.pack(fill=tkinter.Y, side=tkinter.LEFT, expand=tkinter.TRUE)
        self.NEXT_BUTTON.pack(fill=tkinter.BOTH, side=tkinter.LEFT, padx=5, expand=tkinter.TRUE)

    def get_page_no(self):
        return int(self.PAGE_NO.get())

    def set_page_no(self):
        self.PAGE_NO.set(self.TABLE_PAGE_NO.get())


class TableScrollBar(ttk.Scrollbar):
    def __init__(self, _parent: tkinter.Frame, _orient: [tkinter.X, tkinter.Y], _command):
        self._code = 'TableScrollBar'
        self.ORIENT = _orient
        if self.ORIENT == tkinter.X:
            super().__init__(master=_parent, command=_command, orient=tkinter.HORIZONTAL)
            self.pack(side=tkinter.BOTTOM, fill=self.ORIENT)
        else:
            super().__init__(master=_parent, command=_command)
            self.pack(side=tkinter.RIGHT, fill=self.ORIENT)


class Cell(tkinter.Canvas):
    def __init__(self, *args):
        self._code = 'Cell'
        self.TABLE = args[0]
        self.X = args[1]
        self.Y = args[2]
        self.W = args[3]
        self.H = args[4]
        self.COL_ID = args[5]
        self.ROW_ID = args[6]
        self.COL_NO = args[7]
        self.ROW_NO = args[8]
        self.VALUE = args[9]
        self.ICI = args[10]
        self.DC_COMMAND = args[11]
        self.FONT = Font()
        super().__init__(
            self.TABLE,
            background='#9DCAFF',
            borderwidth=0,
            highlightthickness=0
        )
        self.TEXT = self.create_text(0, 0, font=APP_FONT)
        self.coords(self.TEXT, self.W - (self.FONT.measure(self.VALUE) / 2 + (self.W - self.FONT.measure(self.VALUE)) / 2), self.H / 2)
        self.itemconfigure(self.TEXT, text=self.VALUE)
        self.configure(width=self.W, height=self.H)
        self.place(in_=self.TABLE, x=self.X, y=self.Y)
        self.bind('<Button-1>', lambda event, _c=self: self.TABLE.cell_click(_e=event, _c=_c))
        self.bind('<Control-1>', lambda event, _c=self: self.TABLE.cell_ctrl_click(_e=event, _c=_c))
        self.bind('<Shift-1>', lambda event, _c=self: self.TABLE.cell_shift_click(_e=event, _c=_c))
        if self.DC_COMMAND is not None:
            self.bind('<Double-1>', lambda event, ici=self.ICI: self.DC_COMMAND(_ici=ici))


class Table(ttk.Treeview):
    def __init__(self, _app, _parent, _code: str = 'Table'):
        self._code = _code
        self.STYLE = ttk.Style()
        self.STYLE.configure('Treeview', font=APP_FONT)
        self.APP = _app
        self.LOCATION = tkinter.Frame(_parent)
        self.LOCATION_Y = 0
        self.XS = TableScrollBar(self.LOCATION, tkinter.X, lambda *args: self.scrolling_x_table(True, *args))
        self.YS = TableScrollBar(self.LOCATION, tkinter.Y, lambda *args: self.scrolling_y_table(True, *args))
        self.HEAD = None
        self.HEAD_SORT_KEY = None
        self.HEAD_SORT_HOW = 0
        self.HEAD_SORT_STR = None
        self.HEAD_SORT_COMMAND = None
        self.CELL_DOUBLE_CLICK_COMMAND = None
        self.CELL_SELECTIONS = []
        self.CELL_SELECTIONS_RANGE = []
        self.CELL_FONT: Font = Font()
        self.PAGE_NO = None
        self.END_PAGE = None
        super().__init__(
            master=self.LOCATION,
            show='headings',
            xscrollcommand=lambda *args: self.scrolling_x_table(False, *args),
            yscrollcommand=lambda *args: self.scrolling_y_table(False, *args)
        )
        self.CELL_SELECTION_COUNT = tkinter.StringVar(master=self, value='0 개\n복사')
        self.COUNT = {
            'TOT': tkinter.StringVar(master=self, value='0'),
            'NEW': tkinter.StringVar(master=self, value='0'),
            'CHG': tkinter.StringVar(master=self, value='0'),
            'SAM': tkinter.StringVar(master=self, value='0')
        }
        self.XS.config(command=lambda *args: self.scrolling_x_table(True, *args))
        self.YS.config(command=lambda *args: self.scrolling_y_table(True, *args))
        self.pack(fill=tkinter.BOTH, expand=tkinter.TRUE)

    def set_table_paging(self, _flag: bool = False):
        if _flag:
            self.PAGE_NO = tkinter.StringVar(master=self, value='1')
            self.END_PAGE = tkinter.StringVar(master=self, value='1')
        else:
            self.PAGE_NO = None
            self.END_PAGE = None

    def set_table_celling(self, _flag: bool = False):
        if _flag:
            self.configure(selectmode='none')
            self.bind('<Button-1>', lambda event: self.cell_click(_e=event))
            self.bind('<Control-1>', lambda event: self.cell_ctrl_click(_e=event))
            self.bind('<Shift-1>', lambda event: self.cell_shift_click(_e=event))
        else:
            self.configure(selectmode='browse')
            self.unbind('<Button-1>')
            self.unbind('<Control-1>')
            self.unbind('<Shift-1>')

    def set_table_resizing(self, _flag: bool = True):
        if _flag:
            self.APP.bind('<Configure>', self.resizing_table)
        else:
            self.APP.unbind('<Configure>')

    def set_location_y(self, _location_y: int):
        self.LOCATION_Y = _location_y

    def set_count_total(self, _count: int):
        self.COUNT['TOT'].set(str(_count))

    def get_count_total(self):
        return int(self.COUNT['TOT'].get())

    def clear_count(self):
        self.COUNT['TOT'].set('0')
        self.COUNT['NEW'].set('0')
        self.COUNT['CHG'].set('0')
        self.COUNT['SAM'].set('0')

    def insert_row(self, _data, is_tag: bool = False):
        self.COUNT['NEW'].set('0')
        self.COUNT['CHG'].set('0')
        self.COUNT['SAM'].set('0')
        if is_tag:
            for _row_no in range(len(_data)):
                _row_values = list(_data[_row_no])
                for _col_index in range(len(_row_values)):
                    if _row_values[_col_index] is None:
                        _row_values[_col_index] = ''
                if type(_row_values[6]) == datetime.date:
                    _rdd_d_day = (_row_values[6] - datetime.date.today()).days
                    if _rdd_d_day < 0:
                        _row_values[6] = 'OVER) {}'.format(_row_values[6])
                    elif _rdd_d_day <= 10:
                        _row_values[6] = 'D-{}) {}'.format(str(_rdd_d_day).zfill(2), _row_values[6])
                else:
                    _rdd_d_day = 999999999
                _row_status = _row_values.pop(1)
                if _row_status == 2:
                    _row_values.insert(1, '변경')
                    self.COUNT['CHG'].set(str(int(self.COUNT['CHG'].get()) + 1))
                    if _rdd_d_day <= 10:
                        self.insert(parent='', values=_row_values, index=_row_no, tags='changed+rdd_warn')
                    else:
                        self.insert(parent='', values=_row_values, index=_row_no, tags='changed')
                else:
                    if _row_status == 0:
                        _row_values.insert(1, '신규')
                        self.COUNT['NEW'].set(str(int(self.COUNT['NEW'].get()) + 1))
                    else:
                        _row_values.insert(1, '동일')
                        self.COUNT['SAM'].set(str(int(self.COUNT['SAM'].get()) + 1))

                    if _rdd_d_day <= 10:
                        self.insert(parent='', values=_row_values, index=_row_no, tags='rdd_warn')
                    else:
                        self.insert(parent='', values=_row_values, index=_row_no)
            self.tag_configure("changed", background="#9DFFCE")
            self.tag_configure("rdd_warn", foreground='red')
            self.tag_configure("changed+rdd_warn", background="#9DFFCE", foreground='red')
        else:
            for _row in _data:
                if _row[0] != '-':
                    _row = list(_row)
                    for i in range(len(_row)):
                        if _row[i] is None:
                            _row[i] = ''
                    self.insert(parent='', values=_row, index=len(self.get_children()))
                else:
                    self.insert(parent='', values=_row, index=0)

    def delete_row(self):
        for _row_no in self.get_children():
            self.delete(_row_no)

    def scrolling_x_table(self, *args):
        if args[0]:
            self.xview(*args[1:])
        else:
            self.XS.set(first=args[1], last=args[2])
        self.scrolling_cell()

    def scrolling_y_table(self, *args):
        if args[0]:
            self.yview(*args[1:])
        else:
            self.YS.set(first=args[1], last=args[2])
        self.scrolling_cell()

    def resizing_table(self, _e):
        self.LOCATION.place_configure(height = self.APP.get_current_geometry()[1] * 0.94 - self.LOCATION_Y)

    def head_sort(self, _sort_key: str = None):
        if self.HEAD_SORT_KEY == _sort_key:
            if self.HEAD_SORT_HOW == 2:
                self.HEAD_SORT_KEY = None
                self.HEAD_SORT_HOW = 0
            else:
                self.HEAD_SORT_HOW += 1
        else:
            self.HEAD_SORT_KEY = _sort_key
            self.HEAD_SORT_HOW = 1
        self.head_set()

    def head_sort_clear(self):
        self.HEAD_SORT_KEY = None
        self.HEAD_SORT_HOW = 0
        self.head_set()

    def get_head_sort(self):
        _return_value = ''
        if self.HEAD_SORT_KEY is not None:
            if self.HEAD_SORT_HOW == 1:
                _return_value = 'ORDER BY {}'.format(self.HEAD_SORT_KEY)
            elif self.HEAD_SORT_HOW == 2:
                _return_value = 'ORDER BY {} DESC'.format(self.HEAD_SORT_KEY)
        return _return_value

    def set_head_sort_command(self, _sort_command = None):
        self.HEAD_SORT_COMMAND = _sort_command

    def set_cell_double_click_command(self, _double_click_command = None):
        self.CELL_DOUBLE_CLICK_COMMAND = _double_click_command

    def head_set(self, _selected_head: dict = None, _is_change: bool = False):
        if self.HEAD is None:
            self.HEAD = _selected_head
            self['columns'] = list(self.HEAD.keys())
        else:
            if _is_change:
                self.HEAD = _selected_head
                self['columns'] = list(self.HEAD.keys())
        for key in list(self.HEAD.keys()):
            if self.HEAD_SORT_KEY == key:
                if self.HEAD_SORT_HOW == 1:
                    _arrow_text = ' ⬆'
                elif self.HEAD_SORT_HOW == 2:
                    _arrow_text = ' ⬇'
                else:
                    _arrow_text = ''
                self.column(key, anchor=tkinter.CENTER, width=self.HEAD[key]['WIDTH'], stretch=tkinter.NO)
                if self.HEAD_SORT_COMMAND is not None:
                    self.heading(key, text=self.HEAD[key]['TEXT'] + _arrow_text, command=lambda _k=key: self.HEAD_SORT_COMMAND(_head_key=_k))
                else:
                    self.heading(key, text=self.HEAD[key]['TEXT'])
            else:
                self.column(key, anchor=tkinter.CENTER, width=self.HEAD[key]['WIDTH'], stretch=tkinter.NO)
                if self.HEAD_SORT_COMMAND is not None:
                    self.heading(key, text=self.HEAD[key]['TEXT'], command=lambda _k=key: self.HEAD_SORT_COMMAND(_head_key=_k))
                else:
                    self.heading(key, text=self.HEAD[key]['TEXT'])

    def set_end_page(self, _end_page: int):
        if self.PAGE_NO is not None:
            self.END_PAGE.set(str(_end_page))
            if int(self.PAGE_NO.get()) > int(self.END_PAGE.get()):
                self.PAGE_NO.set('1')

    def validate_end_page(self, _view_count: int):
        if self.PAGE_NO is not None:
            if int(self.PAGE_NO.get()) > int(str(divmod(int(self.COUNT['TOT'].get()), _view_count)[0] + (1 if divmod(int(self.COUNT['TOT'].get()), _view_count)[1] > 0 else 0))):
                self.PAGE_NO.set('1')
                return False
            else:
                return True
        else:
            return True

    def get_now_page(self):
        if self.PAGE_NO is not None:
            return self.PAGE_NO.get()
        else:
            return 1

    def validate_page_no(self, _no: int):
        if self.PAGE_NO is not None:
            if 1 <= _no <= int(self.END_PAGE.get()):
                return True
            else:
                return False
        else:
            return True

    def set_prev_page(self):
        if self.PAGE_NO is not None:
            if int(self.PAGE_NO.get()) > 1:
                self.PAGE_NO.set(str(int(self.PAGE_NO.get()) - 1))
            return self.get_now_page()
        else:
            return 1

    def set_move_page(self, _page_no: int):
        if self.PAGE_NO is not None:
            if 1 <= _page_no <= int(self.END_PAGE.get()):
                self.PAGE_NO.set(str(_page_no))
            return self.get_now_page()
        else:
            return 1

    def set_next_page(self):
        if self.PAGE_NO is not None:
            if int(self.PAGE_NO.get()) < int(self.END_PAGE.get()):
                self.PAGE_NO.set(str(int(self.PAGE_NO.get()) + 1))
            return self.get_now_page()
        else:
            return 1

    def cell_clear(self):
        for _cell in self.CELL_SELECTIONS:
            _cell.destroy()
        self.CELL_SELECTIONS.clear()
        self.CELL_SELECTIONS_RANGE.clear()
        self.CELL_SELECTION_COUNT.set('0 개\n복사')

    def cell_click(self, _e=None, _c=None):
        try:
            if _c is None:
                if self.identify_region(_e.x, _e.y) == 'cell':
                    if len(self.identify_row(_e.y)) > 0:
                        for _cell in self.CELL_SELECTIONS:
                            _cell.destroy()
                        self.CELL_SELECTIONS.clear()
                        self.CELL_SELECTIONS_RANGE.clear()
                        _cell = Cell(
                            self,
                            *self.bbox(self.identify_row(_e.y), self.identify_column(_e.x)),
                            self.identify_column(_e.x),
                            self.identify_row(_e.y),
                            int(str(self.identify_column(_e.x)).replace('#', '')) - 1,
                            list(self.get_children()).index(self.identify_row(_e.y)),
                            self.item(self.identify_row(_e.y))['values'][int(str(self.identify_column(_e.x)).replace('#', '')) - 1],
                            self.item(self.identify_row(_e.y))['values'][0],
                            self.CELL_DOUBLE_CLICK_COMMAND
                        )
                        if len(self.CELL_SELECTIONS_RANGE) < 1:
                            self.CELL_SELECTIONS_RANGE.append({
                                'START': _cell,
                                'CLOSE': None
                            })
                        else:
                            if self.CELL_SELECTIONS_RANGE[-1]['CLOSE'] is None:
                                self.CELL_SELECTIONS_RANGE[-1]['START'] = _cell
                            else:
                                self.CELL_SELECTIONS_RANGE.append({
                                    'START': _cell,
                                    'CLOSE': None
                                })
                        self.CELL_SELECTIONS.append(_cell)
            else:
                _ = []
                for _cell in self.CELL_SELECTIONS:
                    if _c != _cell:
                        _cell.destroy()
                        _.append(_cell)
                for _destroyed in _:
                    self.CELL_SELECTIONS.remove(_destroyed)
                self.CELL_SELECTIONS_RANGE.clear()
                self.CELL_SELECTIONS_RANGE.append(
                    {
                        'START': _c,
                        'CLOSE': None
                    }
                )
        except ValueError:
            print(sys.exc_info()[0])
            print(sys.exc_info()[1])
            print(sys.exc_info()[2].tb_lineno)
        finally:
            self.CELL_SELECTION_COUNT.set(str(len(self.CELL_SELECTIONS)) + ' 개\n복사')

    def cell_ctrl_click(self, _e=None, _c=None):
        try:
            if _c is None:
                if self.identify_region(_e.x, _e.y) == 'cell':
                    if len(self.identify_row(_e.y)) > 0:
                        _cell = Cell(
                            self,
                            *self.bbox(self.identify_row(_e.y), self.identify_column(_e.x)),
                            self.identify_column(_e.x),
                            self.identify_row(_e.y),
                            int(str(self.identify_column(_e.x)).replace('#', '')) - 1,
                            list(self.get_children()).index(self.identify_row(_e.y)),
                            self.item(self.identify_row(_e.y))['values'][int(str(self.identify_column(_e.x)).replace('#', '')) - 1],
                            self.item(self.identify_row(_e.y))['values'][0],
                            self.CELL_DOUBLE_CLICK_COMMAND
                        )
                        if len(self.CELL_SELECTIONS_RANGE) < 1:
                            self.CELL_SELECTIONS_RANGE.append({
                                'START': _cell,
                                'CLOSE': None
                            })
                        else:
                            if self.CELL_SELECTIONS_RANGE[-1]['CLOSE'] is None:
                                self.CELL_SELECTIONS_RANGE[-1]['START'] = _cell
                            else:
                                self.CELL_SELECTIONS_RANGE.append({
                                    'START': _cell,
                                    'CLOSE': None
                                })
                        self.CELL_SELECTIONS.append(_cell)
            else:
                _c.destroy()
                self.CELL_SELECTIONS.remove(_c)
        except ValueError:
            print(sys.exc_info()[0])
            print(sys.exc_info()[1])
            print(sys.exc_info()[2].tb_lineno)
        finally:
            self.CELL_SELECTION_COUNT.set(str(len(self.CELL_SELECTIONS)) + ' 개\n복사')

    def cell_shift_click(self, _e=None, _c=None):
        try:
            if self.identify_region(_e.x, _e.y) == 'cell':
                if len(self.identify_row(_e.y)) > 0 or _c is not None:
                    if len(self.CELL_SELECTIONS) < 1:
                        self.cell_click(_e)
                    else:
                        _start_cell = self.CELL_SELECTIONS_RANGE[-1]['START']
                        _close_cell = self.CELL_SELECTIONS_RANGE[-1]['CLOSE']
                        if _close_cell is not None:
                            if _start_cell.ROW_NO < _close_cell.ROW_NO:
                                _f_row = _start_cell.ROW_NO
                                _b_row = _close_cell.ROW_NO + 1
                            elif _start_cell.ROW_NO == _close_cell.ROW_NO:
                                _f_row = _start_cell.ROW_NO
                                _b_row = _start_cell.ROW_NO + 1
                            else:
                                _f_row = _close_cell.ROW_NO
                                _b_row = _start_cell.ROW_NO + 1
                            if _start_cell.COL_NO < _close_cell.COL_NO:
                                _f_col = _start_cell.COL_NO
                                _b_col = _close_cell.COL_NO + 1
                            elif _start_cell.COL_NO == _close_cell.COL_NO:
                                _f_col = _start_cell.COL_NO
                                _b_col = _start_cell.COL_NO + 1
                            else:
                                _f_col = _close_cell.COL_NO
                                _b_col = _start_cell.COL_NO + 1
                            _ = []
                            for _row in range(_f_row, _b_row):
                                for _col in range(_f_col, _b_col):
                                    for _cell in self.CELL_SELECTIONS:
                                        if _cell.ROW_NO == _row and _cell.COL_NO == _col:
                                            if not (_start_cell.ROW_NO == _row and _start_cell.COL_NO == _col):
                                                if _c is not None:
                                                    if not (_c.ROW_NO == _row and _c.COL_NO == _col):
                                                        _cell.destroy()
                                                        _.append(_cell)
                                                else:
                                                    _cell.destroy()
                                                    _.append(_cell)
                            for _destroyed in _:
                                self.CELL_SELECTIONS.remove(_destroyed)
                        if _c is not None:
                            _close_cell = _c
                            _close_x, _close_y, _close_w, _close_h = _close_cell.X, _close_cell.Y, _close_cell.W, _close_cell.H
                        else:
                            _close_cell = Cell(
                                self,
                                *self.bbox(self.identify_row(_e.y), self.identify_column(_e.x)),
                                self.identify_column(_e.x),
                                self.identify_row(_e.y),
                                int(str(self.identify_column(_e.x)).replace('#', '')) - 1,
                                list(self.get_children()).index(self.identify_row(_e.y)),
                                self.item(self.identify_row(_e.y))['values'][int(str(self.identify_column(_e.x)).replace('#', '')) - 1],
                                self.item(self.identify_row(_e.y))['values'][0],
                                self.CELL_DOUBLE_CLICK_COMMAND
                            )
                            _close_x, _close_y, _close_w, _close_h = _close_cell.X, _close_cell.Y, _close_cell.W, _close_cell.H
                        if _start_cell.ROW_NO < _close_cell.ROW_NO:
                            _f_row = _start_cell.ROW_NO
                            _b_row = _close_cell.ROW_NO + 1
                        elif _start_cell.ROW_NO == _close_cell.ROW_NO:
                            _f_row = _start_cell.ROW_NO
                            _b_row = _start_cell.ROW_NO + 1
                        else:
                            _f_row = _close_cell.ROW_NO
                            _b_row = _start_cell.ROW_NO + 1
                        if _start_cell.COL_NO < _close_cell.COL_NO:
                            _f_col = _start_cell.COL_NO
                            _b_col = _close_cell.COL_NO + 1
                        elif _start_cell.COL_NO == _close_cell.COL_NO:
                            _f_col = _start_cell.COL_NO
                            _b_col = _start_cell.COL_NO + 1
                        else:
                            _f_col = _close_cell.COL_NO
                            _b_col = _start_cell.COL_NO + 1
                        for _row in range(_f_row, _b_row):
                            for _col in range(_f_col, _b_col):
                                if _start_cell.ROW_NO == _row and _start_cell.COL_NO == _col:
                                    self.CELL_SELECTIONS_RANGE[-1]['START'] = _start_cell
                                else:
                                    if _close_cell.ROW_NO == _row and _close_cell.COL_NO == _col:
                                        self.CELL_SELECTIONS_RANGE[-1]['CLOSE'] = _close_cell
                                        self.CELL_SELECTIONS.append(_close_cell)
                                    else:
                                        try:
                                            _cx, _cy, _cw, _ch = self.bbox(list(self.get_children())[_row], '#{}'.format(_col + 1))
                                        except ValueError:
                                            _cx, _cy, _cw, _ch = -9999, -9999, [_value['WIDTH'] for _value in self.HEAD.values()][_col], 20
                                        _cell = Cell(
                                            self,
                                            _cx,
                                            _cy,
                                            _cw,
                                            _ch,
                                            '#{}'.format(_col + 1),
                                            list(self.get_children())[_row],
                                            _col,
                                            _row,
                                            self.item(list(self.get_children())[_row])['values'][_col],
                                            self.item(list(self.get_children())[_row])['values'][0],
                                            self.CELL_DOUBLE_CLICK_COMMAND
                                        )
                                        self.CELL_SELECTIONS.append(_cell)
        except ValueError:
            print(sys.exc_info()[0])
            print(sys.exc_info()[1])
            print(sys.exc_info()[2].tb_lineno)
        finally:
            self.CELL_SELECTION_COUNT.set(str(len(self.CELL_SELECTIONS)) + ' 개\n복사')

    def scrolling_cell(self):
        for _cell in self.CELL_SELECTIONS:
            if _cell.winfo_exists():
                try:
                    _x, _y, _, _ = self.bbox(_cell.ROW_ID, _cell.COL_ID)
                    if _x != _cell.X:
                        _cell.X = _x
                        _cell.place_configure(in_=self, x=_x, y=_cell.Y)
                    if _y != _cell.Y:
                        _cell.Y = _y
                        _cell.place_configure(in_=self, x=_cell.X, y=_y)
                except ValueError:
                    _cell.X = -9999
                    _cell.Y = -9999
                    _cell.place_forget()

    def selection_copy(self, _event=None):
        _clipboard_string = ''
        try:
            _clipboard_data = []
            _min_row = 999999999
            _max_row = 0
            _min_col = 999999999
            _max_col = 0
            for _cell in self.CELL_SELECTIONS:
                if _min_row > _cell.ROW_NO:
                    _min_row = _cell.ROW_NO
                if _min_col > _cell.COL_NO:
                    _min_col = _cell.COL_NO
                if _max_row < _cell.ROW_NO:
                    _max_row = _cell.ROW_NO
                if _max_col < _cell.COL_NO:
                    _max_col = _cell.COL_NO
            for _row in range(_min_row, _max_row + 1):
                _ = []
                for _col in range(_min_col, _max_col + 1):
                    for _cell in self.CELL_SELECTIONS:
                        if _cell.ROW_NO == _row and _cell.COL_NO == _col:
                            _cell_value = str(_cell.VALUE).replace('OVER) ', '')
                            for d_no in range(0, 11):
                                _cell_value =_cell_value.replace('D-{}) '.format(str(d_no).zfill(2)), '')
                            _.append(_cell_value)
                            break
                    else:
                        _.append('')
                _clipboard_data.append(_)
            for _row_data in _clipboard_data:
                for _col_data in _row_data:
                    _clipboard_string += '{}\t'.format(_col_data if len(_col_data) != 0 else ' ')
                _clipboard_string += '\n'
            return True, _clipboard_string
        except:
            return False, {
                'ECD': self._code + '_SelectionCopy',
                'CLS': sys.exc_info()[0],
                'DES': sys.exc_info()[1],
                'LNO': sys.exc_info()[2].tb_lineno,
                'DTA': [len(self.CELL_SELECTIONS), _clipboard_string]
            }
