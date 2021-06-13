# -*- coding: utf-8 -*-

from _version import __version__
from checkboxtreeview import CheckboxTreeview
from calendar import month_name
from datetime import date, datetime
from decimal import Decimal
from label_grid import LabelGrid
from multiselect import MultiselectMenu
from tkcalendar import DateEntry
from tkinter import ttk, messagebox
from tkinter.filedialog import askopenfile
from shutil import copy, copy2
from pathlib import Path
from tkinter import filedialog as fd
import tkinter.font as tkFont
from tkHyperlinkManager import HyperlinkManager
from math import floor
from xl import export_to_excel
import datetime as dt
import locale
import os, zlib
import tkinter as tk
import ast

# example of subsription and default recipient
EMAIL_TO = b'\xd0\xa4\xd0\xbe\xd0\xb7\xd0\xb7\xd0\xb8|\
\xd0\x9b\xd0\xbe\xd0\xb3\xd0\xb8\xd1\x81\xd1\x82\xd0\xb8\xd0\xba\xd0\xb0|\
\xd0\x90\xd0\xbd\xd0\xb0\xd0\xbb\xd0\xb8\xd1\x82\xd0\xb8\xd0\xba\xd0\xb8'.decode()
# example of path to independent report
TEMPLATE_PATH = zlib.decompress(b'x\x9c\x8b\x89I\xcb\xaf\xaa\xaa\xd4\xcbI\xcc'
                                b'\x8bq\xc9O.\xcdM\xcd+)\x8e\xf1\xc8\xcfI\xc9'
                                b'\xccK\x8fqI-H,*\x81\x88\x05g\xe6\x14\xe4\xc7'
                                b'\\\x98}a\xdf\x85\xcd\x17v\\l\xbc\xd8ta\xc7\x85]'
                                b'\x176\xc4\xbb\xbb\x06\xf9\xba\x06\xc7\x04\xa4'
                                b'\x16\xa5\xe5\x17\xe5\xa6\x16\xc5\x94\xa4&g\xc4'
                                b'\x04\xa5&\x17\x95f\x96\x00M\x01\x00\xa74-'
                                b'\x18').decode()

UPLOAD_PATH = zlib.decompress(b"x\x9c\x8b\x89I\xcb\xaf\xaa\xaa\xd4\xcbI\xcc"
                              b"\x8bq\xc9O.\xcdM\xcd+)\x8e\xf1\xc8\xcfI\xc9"
                              b"\xccK\x8fqI-H,*\x81\x88\xf9\xe4\xa7g\x16"
                              b"\x97df'\xc6\xb8e\xe6\xc5\\\x98x\xb1\xef\xc2"
                              b"\x96\x0b\xdb.l\xbd\xd8\x14\x13Z\x90\x93\x9f"
                              b"\x98\x02\x00\xa3\x8c!\xb1").decode()


class PaymentsError(Exception):
    """Base class for exceptions in this module."""
    pass


class IncorrectFloatError(PaymentsError):
    """ Exception raised if sum is not converted to float.

    Attributes:
        expression - input expression in which the error occurred;
        message - explanation of the error.
    """

    def __init__(self, expression, message='Введена некорректная сумма'):
        self.expression = expression
        self.message = message
        super().__init__(self.expression, self.message)


class MonthFilterError(PaymentsError):
    """ Exception raised if month don't chosen in filter.

    Attributes:
        expression - input expression in which the error occurred;
        message - explanation of the error.
    """

    def __init__(self, expression, message='Не выбран месяц'):
        self.expression = expression
        self.message = message
        super().__init__(self.expression, self.message)


class AccessError(tk.Tk):
    """ Raise an error when user doesn't have permission to work with app.
    """

    def __init__(self):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showerror(
            'Ошибка доступа',
            'Нет прав для работы с приложением.\n'
            'Для получения прав обратитесь на рассылку ' + EMAIL_TO
        )
        self.destroy()


class LoginError(tk.Tk):
    """ Raise an error when user doesn't have permission to work with db.
    """

    def __init__(self):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showerror(
            'Ошибка подключения',
            'Нет прав для работы с сервером.\n'
            'Обратитесь на рассылку ' + EMAIL_TO
        )
        self.destroy()


class NetworkError(tk.Tk):
    """ Raise a message about network error.
    """

    def __init__(self):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showerror(
            'Ошибка cети',
            'Возникла общая ошибка сети.\nПерезапустите приложение'
        )
        self.destroy()


class RestartRequiredAfterUpdateError(tk.Tk):
    """ Raise a message about restart needed after update.
    """

    def __init__(self):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showinfo(
            'Необходима перезагрузка',
            'Выполнено критическое обновление.\nПерезапустите приложение'
        )
        self.destroy()


class UnexpectedError(tk.Tk):
    """ Raise a message when an unexpected exception occurs.
    """

    def __init__(self, *args):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showinfo(
            'Непредвиденное исключение',
            'Возникло непредвиденное исключение\n' + '\n'.join(map(str, args))
        )
        self.destroy()


class StringSumVar(tk.StringVar):
    """ Contains function that returns var formatted in a such way, that
        it can be converted into a float without an error.
    """

    def get_float_form(self, *args, **kwargs):
        return super().get(*args, **kwargs).replace(' ', '').replace(',', '.')


class RecruitingApp(tk.Tk):
    def __init__(self, **kwargs):
        super().__init__()
        self.title('Заявки на персонал')
        self.iconbitmap('resources/file.ico')
        # store the state of PreviewForm
        self.state_PreviewForm = 'normal'
        # geometry_storage {Framename:(width, height)}
        self._geometry = {'PreviewForm': (1200, 550),
                          'CreateForm': (480, 440),
                          'UpdateForm': (480, 300)}
        # Virtual event for creating request
        self.event_add("<<create>>", "<Control-S>", "<Control-s>",
                       "<Control-Ucircumflex>", "<Control-ucircumflex>",
                       "<Control-twosuperior>", "<Control-threesuperior>",
                       "<KeyPress-F5>")
        self.bind_all("<Key>", self._onKeyRelease, '+')
        self.bind("<<create>>", self._create_request)
        self.active_frame = None
        # handle the window close event
        self.protocol("WM_DELETE_WINDOW", self._quit)
        # hide until all frames have been created
        self.withdraw()
        # To import months names in cyrillic
        locale.setlocale(locale.LC_ALL, 'RU')
        # Customize header style (used in PreviewForm)
        style = ttk.Style()
        try:
            style.element_create("HeaderStyle.Treeheading.border", "from",
                                 "default")
            style.layout("HeaderStyle.Treeview.Heading", [
                ("HeaderStyle.Treeheading.cell", {'sticky': 'nswe'}),
                ("HeaderStyle.Treeheading.border",
                 {'sticky': 'nswe', 'children': [
                     ("HeaderStyle.Treeheading.padding",
                      {'sticky': 'nswe', 'children': [
                          ("HeaderStyle.Treeheading.image",
                           {'side': 'right', 'sticky': ''}),
                          ("HeaderStyle.Treeheading.text", {'sticky': 'we'})
                      ]})
                 ]}),
            ])
            style.configure("HeaderStyle.Treeview.Heading",
                            background="#f1f1f1", foreground="black",
                            relief='groove', font=('Arial', 8))
            style.map("HeaderStyle.Treeview.Heading",
                      relief=[('active', 'sunken'), ('pressed', 'flat')])

            style.map('ButtonGreen.TButton')
            style.configure('ButtonGreen.TButton', foreground='green')

            style.map('ButtonRed.TButton')
            style.configure('ButtonRed.TButton', foreground='red')

            style.configure("TMenubutton", background='white')
        except tk.TclError:
            # if during debug previous tk wasn't destroyed
            # and style remains modified
            pass

        # the container is where we'll stack a bunch of frames
        # the one we want to be visible will be raised above others
        container = tk.Frame(self)
        container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self._frames = {}
        for F in (PreviewForm, CreateForm, UpdateForm):
            frame_name = F.__name__
            frame = F(parent=container, controller=self, **kwargs)
            self._frames[frame_name] = frame
            # put all of them in the same location
            frame.grid(row=0, column=0, sticky='nsew')

        self._show_frame('PreviewForm')
        # restore after withdraw
        self.deiconify()

    def _center_window(self, w, h):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        start_x = int((screen_width / 2) - (w / 2))
        start_y = int((screen_height / 2) - (h * 0.55))

        self.geometry('{}x{}+{}+{}'.format(w, h, start_x, start_y))

    def _fill_UpdateForm(self, **kwargs):
        """ Control function to transfer data from Preview- to CreateForm. """
        id = kwargs['ID']
        internalID = kwargs['Номер заявки']
        officeName = kwargs['Офис']
        departmentName = kwargs['Департамент']
        responsibleUser = kwargs['Ответственный от HR']
        statusID = kwargs['StatusID']
        fileCV = kwargs['Файл резюме']
        frame = self._frames['UpdateForm']
        frame._fill_from_UpdateForm(id, internalID, officeName, departmentName,
                                    responsibleUser, statusID, fileCV)

    def _onKeyRelease(*args):
        event = args[1]
        # check if Ctrl pressed
        if not (event.state == 12 or event.state == 14):
            return
        if event.keycode == 88 and event.keysym.lower() != 'x':
            event.widget.event_generate("<<Cut>>")
        elif event.keycode == 86 and event.keysym.lower() != 'v':
            event.widget.event_generate("<<Paste>>")
        elif event.keycode == 67 and event.keysym.lower() != 'c':
            event.widget.event_generate("<<Copy>>")

    def _show_frame(self, frame_name):
        """ Show a frame for the given frame name. """
        if frame_name == 'CreateForm' or frame_name == 'UpdateForm':
            # since we have only two forms, when we activating CreateForm
            # we know by exception that PreviewForm is active
            self.state_PreviewForm = self.state()
            self.state('normal')
            self.resizable(width=False, height=False)
        else:
            self.state(self.state_PreviewForm)
            self.resizable(width=True, height=True)
        frame = self._frames[frame_name]
        frame.tkraise()
        self._center_window(*(self._geometry[frame_name]))
        # Make sure active_frame changes in case of network error
        try:
            if frame_name in ('PreviewForm'):
                frame._resize_columns()
                frame._refresh()
                # Clear form in CreateFrom and UpdateForm by autofill form
                self._frames['CreateForm']._clear(0)
                self._frames['UpdateForm']._clear()
        finally:
            self.active_frame = frame_name

    def _create_request(self, event):
        """
        Creates request when hotkey is pressed if active_frame is CreateForm.
        """
        if self.active_frame == 'CreateForm':
            self._frames[self.active_frame]._create_request()

    def _quit(self):
        if self.active_frame != 'PreviewForm':
            self._show_frame('PreviewForm')
        elif messagebox.askokcancel("Выход", "Выйти из приложения?"):
            self.destroy()


class RecruitingFrame(tk.Frame):
    def __init__(self, parent, controller, connection, user_info, office):
        super().__init__(parent)
        self.parent = parent
        self.controller = controller
        self.conn = connection
        self.office = {}
        if isinstance(self, PreviewForm):
            self.officeID, self.office = zip(*[(None, 'Все'), ] + office)
        self.user_info = user_info
        # Often used info
        self.userID = user_info.UserID
        self.userOffice = user_info.OfficeName
        self.userDepartment = user_info.DepartmentName
        self.userPosition = user_info.Position

    def _add_user_label(self, parent):
        """ Adds user name in top right corner. """
        user_label = tk.Label(parent, text='Пользователь: ' +
                                           self.user_info.ShortUserName + '  Версия ' + __version__,
                              font=('Arial', 8))
        user_label.pack(side=tk.RIGHT, anchor=tk.NE)

    def get_officeID(self, office):
        return self.office[office][0]


class CreateForm(RecruitingFrame):
    def __init__(self, parent, controller, connection, user_info, office, **kwargs):
        super().__init__(parent, controller, connection, user_info, office)
        self.uploaded_filename = str()
        # Top Frame with description and user name
        top = tk.Frame(self, name='top_cf', padx=5)
        self.main_label = tk.Label(top,
                                   text='Форма создания заявки на поиск персонала',
                                   padx=10, pady=5, font=('Calibri', 11, 'bold'))
        self._top_pack()

        # First Fill Frame
        row1_cf = tk.Text(self, padx=18, height=3, relief=tk.FLAT, bg='#f1f1f1')
        row1_cf.insert(tk.INSERT, 'Подразделение инициатора:')
        row1_cf.insert(tk.INSERT, str('\n' + self.userOffice))
        row1_cf.insert(tk.INSERT, str('\n' + self.userDepartment))
        row1_cf.tag_add('title', 1.0, '1.end')
        row1_cf.tag_add('style', 2.0, '2.end')
        row1_cf.tag_add('style', 3.0, '3.end')
        row1_cf.tag_config('title', font=("Calibri", 10, 'bold'),
                           justify=tk.LEFT)
        row1_cf.tag_config('style', font=("Calibri", 10, 'normal'),
                           justify=tk.LEFT)
        row1_cf.configure(state="disabled")

        self._row1_pack()

        # Second Fill Frame
        row2_cf = tk.Frame(self, name='row2_cf', padx=10)
        self.separator = ttk.Separator(row2_cf, orient='horizontal')



        self._row2_pack()

        # Third Fill Frame
        row3_cf = tk.Frame(self, name='row3_cf', padx=10)
        self.requirements_label = tk.Label(row3_cf,
                                      text='1. Откройте и заполните файл требований:',
                                      padx=8)
        bt_open_file = ttk.Button(row3_cf, text="Открыть", width=20,
                                  command=self.open_file_requirements)
        bt_open_file.pack(side=tk.RIGHT, padx=15, pady=0)



        self._row3_pack()

        # Fourth Fill Frame
        row4_cf = tk.Frame(self, name='row4_cf', padx=10)
        self.attach_label = tk.Label(row4_cf,
                                   text='2. Прикрепите файл требований:',
                                   padx=8)
        self.upload_btn_text = tk.StringVar()
        bt_upload = ttk.Button(row4_cf, textvariable=self.upload_btn_text,
                               width=20,
                               command=self._upload_requirements,
                               style='ButtonGreen.TButton')
        self.upload_btn_text.set("Выбрать файл")
        bt_upload.pack(side=tk.RIGHT, padx=15, pady=0)

        self._row4_pack()

        # Fifth Fill Frame
        row5_cf = tk.Frame(self, name='row5_cf', padx=10)
        self.candidatePositionLabel = tk.Label(row5_cf,
                                               text='Название должности кандидата:',
                                               padx=7)
        self.candidatePositionEntry = tk.Entry(row5_cf, width=40)

        self._row5_pack()

        # Six Fill Frame
        row6_cf = tk.Frame(self, name='row6_cf', padx=10)
        self.plannedClosingDateLabel = tk.Label(row6_cf,
                                                text='Плановая дата закрытия заявки:',
                                                padx=7)
        self.plannedClosingDate = tk.StringVar()
        self.plannedClosingDateWidget = DateEntry(row6_cf, width=16,
                                                  state='readonly',
                                                  textvariable=self.plannedClosingDate,
                                                  font=('Arial', 9),
                                                  selectmode='day',
                                                  borderwidth=2,
                                                  locale='ru_RU')


        # Text Frame
        text_cf = ttk.LabelFrame(self, text=' Комментарий к заявке ',
                                 name='text_cf')

        self.customFont = tkFont.Font(family="Arial", size=10)
        self.desc_text = tk.Text(text_cf,
                                 font=self.customFont)  # input and output box
        self.desc_text.configure(width=100)
        self.desc_text.pack(in_=text_cf, expand=True)

        self._row6_pack()

        # Bottom Frame with buttons
        bottom_cf = tk.Frame(self, name='bottom_cf')

        bt3 = ttk.Button(bottom_cf, text="Назад", width=10,
                         command=lambda: controller._show_frame('PreviewForm'))
        bt3.pack(side=tk.RIGHT, padx=15, pady=10)

        bt2 = ttk.Button(bottom_cf, text="Очистить", width=10,
                         command=lambda: self._clear(1), style='ButtonRed.TButton')
        bt2.pack(side=tk.RIGHT, padx=0, pady=0)

        bt1 = ttk.Button(bottom_cf, text="Создать заявку", width=15,
                         command=self._create_request,
                         style='ButtonGreen.TButton')
        bt1.pack(side=tk.RIGHT, padx=15, pady=10)

        # Pack frames
        top.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        bottom_cf.pack(side=tk.BOTTOM, fill=tk.X)
        row1_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        row2_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        row3_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        row4_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        row5_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        row6_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        text_cf.pack(side=tk.TOP, fill=tk.X, expand=True, padx=15, pady=15)

    def open_file_requirements(self):
        pathToFile = 'resources\\Требования.docx'
        return os.startfile(pathToFile)

    def _upload_requirements(self):
        filename = fd.askopenfilename()
        if filename:
            # Rename file while uploading
            now = str(datetime.now())[:19]
            now = now.replace(":", "_")
            now = now.replace(" ", "_")
            new_filename = "Требования_" + now + ".docx"
            distPath = UPLOAD_PATH + "\\" + new_filename
            copy(filename, distPath)
            path = Path(distPath)
            self.uploaded_filename = path.name
            self.upload_btn_text.set("Файл добавлен")

    def _remove_uploaded_file(self):
        os.remove(UPLOAD_PATH + '\\' + self.uploaded_filename)

    def _clear(self, param):
        self.candidatePositionEntry.configure(state="normal")
        self.candidatePositionEntry.delete(0, tk.END)
        self.desc_text.delete("1.0", tk.END)
        self.upload_btn_text.set("Выбрать файл")
        self.upload_filename = str()
        self.plannedClosingDateWidget.set_date(datetime.now())
        if self.uploaded_filename and param == 1:
            self._remove_uploaded_file()

    def _fill_from_PreviewForm(self, office, num_main_contract_entry,
                               date_main_contract_start, date_main_contract_end
                               , contragent, responsible, okpo):
        """ When button "Добавить из договора" from PreviewForm is activated,
        fill some fields taken from choosed in PreviewForm request.
        """
        self.office_current.set(office)
        self.responsible_box.set(responsible)
        self.responsible_box.configure(state="readonly")
        self.num_main_contract_entry.delete(0, tk.END)
        self.num_main_contract_entry.insert(0, num_main_contract_entry)
        self.num_main_contract_entry.configure(state="readonly")
        self.date_main_contract_start.set_date(
            self._convert_str_date(date_main_contract_start))
        self.date_main_contract_start.configure(state="readonly")
        self.date_main_contract_end.set_date(
            self._convert_str_date(date_main_contract_end))
        self.date_main_contract_end.configure(state="readonly")
        self.contragent_entry.delete(0, tk.END)
        self.contragent_entry.insert(0, contragent)
        self.contragent_entry.configure(state="readonly")
        self.okpo_entry.insert(0, okpo)
        self.square_cost.set('0,00')

    def _convert_date(self, date, output=None):
        """ Take date and convert it into output format.
            If output is None datetime object is returned.

            date: str in format '%d[./]%m[./]%y' or '%d[./]%m[./]%Y'.
            output: str or None, output format.
        """
        date = date.replace('/', '.')
        try:
            dat = datetime.strptime(date, '%d.%m.%y')
        except ValueError:
            dat = datetime.strptime(date, '%d.%m.%Y')
        if output:
            return dat.strftime(output)
        return dat

    def _create_request(self):
        messagetitle = 'Создание заявки'
        is_validated = self._validate_request_creation(messagetitle)
        if not is_validated:
            return

        request = {'userID': self.userID,
                   'positionName': self.candidatePositionEntry.get(),
                   'plannedDate': self._convert_date(
                       self.plannedClosingDateWidget.get()),
                   'fileRequirements': self.uploaded_filename,
                   'commentText': self.desc_text.get("1.0", tk.END).strip()

                   }
        created_success = self.conn.create_request(**request)
        if created_success == 1:
            messagebox.showinfo(
                messagetitle, 'Заявка на поиск персонала создана'
            )
            self._clear(0)
            self.controller._show_frame('PreviewForm')
        else:
            self._remove_uploaded_file()
            messagebox.showerror(
                messagetitle, 'Произошла ошибка при добавлении заявки'
            )

    def _convert_str_date(self, date):
        """ Take str and convert it into date format.
            date: str in format '%d[./]%m[./]%y' or '%d[./]%m[./]%Y'.
        """
        date_time_str = date
        date_time_obj = dt.datetime.strptime(date_time_str, '%Y-%m-%d')
        return date_time_obj.date()

    def _row1_pack(self):
        pass

    def _row2_pack(self):
        self.separator.pack(fill='x')

    def _row3_pack(self):
        self.requirements_label.pack(side=tk.LEFT, padx=0)

    def _row4_pack(self):
        self.attach_label.pack(side=tk.LEFT, padx=0)

    def _row5_pack(self):
        self.candidatePositionLabel.pack(side=tk.LEFT)
        self.candidatePositionEntry.pack(side=tk.LEFT, padx=2)

    def _row6_pack(self):
        self.plannedClosingDateLabel.pack(side=tk.LEFT)
        self.plannedClosingDateWidget.pack(side=tk.LEFT, padx=3)

    def _top_pack(self):
        self.main_label.pack(side=tk.TOP, expand=False, anchor=tk.NW)

    def _validate_request_creation(self, messagetitle):
        """ Check if all fields are filled properly. """
        if not self.uploaded_filename:
            messagebox.showerror(
                messagetitle, 'Вы не загрузили файл требований'
            )
            return False
        if not self.candidatePositionEntry.get():
            messagebox.showerror(
                messagetitle, 'Не указана должность вакансии'
            )
            return False
        return True


class UpdateForm(RecruitingFrame):
    def __init__(self, parent, controller, connection, user_info, office,
                 responsible_all, **kwargs):
        super().__init__(parent, controller, connection, user_info, office)
        self.responsibleID, self.responsible = zip(*[(0, 'Не назначен'), ] + responsible_all)
        self.customStatusID, self.customStatusName = zip(*[(0, 'Не выбрано'),
                                                           (4, 'Верифицировать'),
                                                           (2, 'Вернуть в работу')])
        self.UserID = self.user_info.UserID
        self.isSuperHR = self.user_info.isSuperHR
        self.filenameCV = str()
        # Top Frame with description and user name
        top = tk.Frame(self, name='top_cf', padx=5)
        self.main_label = tk.Label(top,
                                   text='Управление заявкой на поиск персонала',
                                   padx=10, font=('Calibri', 11, 'bold'))
        self._top_pack()

        # First Fill Frame
        row1_cf = tk.Frame(self, name='row1_cf', padx=15)
        self.request_info_text = tk.StringVar()
        self.request_info_label = tk.Label(row1_cf,
                                           textvariable=self.request_info_text,
                                           padx=7, justify=tk.LEFT, font=('Calibri', 10))

        self._row1_pack()

        # Second Fill Frame
        row2_cf = tk.Frame(self, name='row2_cf', padx=15)
        self.separator = ttk.Separator(row2_cf, orient='horizontal')

        self._row2_pack()

        # Third Fill Frame
        row3_cf = tk.Frame(self, name='row3_cf', padx=15)
        self.responsible_label = tk.Label(row3_cf, text='Выбрать ответственного:', padx=7)
        self.responsible_box = ttk.Combobox(row3_cf, width=20,
                                            state='readonly')
        self.responsible_box['values'] = self.responsible


        self._row3_pack()

        # Fourth Fill Frame
        row4_cf = tk.Frame(self, name='row4_cf', padx=15)
        self.attach_label = tk.Label(row4_cf,
                                     text='Резюме согласованного кандидата:',
                                     padx=8)
        self.upload_btn_text = tk.StringVar()
        self.bt_upload = ttk.Button(row4_cf, textvariable=self.upload_btn_text,
                                    width=23,
                                    command=self._upload_cv,
                                    style='ButtonGreen.TButton',
                                    state=tk.NORMAL)
        self.upload_btn_text.set("Выбрать файл")
        self.bt_upload.pack(side=tk.RIGHT, padx=15, pady=0)

        self._row4_pack()

        # Fifth Fill Frame
        row5_cf = tk.Frame(self, name='row5_cf', padx=15)
        self.status_label = tk.Label(row5_cf,
                                          text='Верифицировать или вернуть в работу:',
                                          padx=7)
        self.status_box = ttk.Combobox(row5_cf, width=20,
                                            state='readonly')
        self.status_box['values'] = self.customStatusName


        self._row5_pack()

        # Six Fill Frame
        row6_cf = tk.Frame(self, name='row6_cf', padx=15)


        self._row6_pack()

        # Bottom Frame with buttons
        bottom_cf = tk.Frame(self, name='bottom_cf')

        bt3 = ttk.Button(bottom_cf, text="Назад", width=10,
                         command=lambda: controller._show_frame('PreviewForm'))
        bt3.pack(side=tk.RIGHT, padx=15, pady=10)

        bt1 = ttk.Button(bottom_cf, text="Обновить", width=10,
                         command=self._update_request,
                         style='ButtonGreen.TButton')
        bt1.pack(side=tk.RIGHT, padx=15, pady=10)

        # Pack frames
        top.pack(side=tk.TOP, fill=tk.BOTH)
        bottom_cf.pack(side=tk.BOTTOM, fill=tk.X)
        row1_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        row2_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        row3_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        row4_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        row5_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        row6_cf.pack(side=tk.TOP, fill=tk.X, pady=5)


    def _upload_cv(self):
        filename = fd.askopenfilename()
        if filename:
            # Rename file while it uploading
            file = Path(filename).name
            now = str(datetime.now())[:19]
            now = now.replace(":", "_")
            now = now.replace(" ", "_")
            new_filename = file.replace(".", '_' + now + '.')
            distPath = UPLOAD_PATH + "\\" + new_filename
            copy(filename, distPath)
            path = Path(distPath)
            self.upload_filename = path.name
            self.upload_btn_text.set("Файл добавлен")

    def _remove_uploaded_file(self):
        os.remove(UPLOAD_PATH + '\\' + self.filenameCV)


    def _clear(self):
        self.responsible_box.configure(state="readonly")
        self.upload_filename = str()

    def _fill_from_UpdateForm(self, id, internalID, officeName,
                              departmentName, responsibleUser, statusID, fileCV):
        """ When button "Управление заявкой" from PreviewForm is activated,
        fill some fields taken from choosed in PreviewForm request.
        """
        self.request_id = id
        self.responsible_box.set(responsibleUser)
        self.filenameCV = fileCV
        if not self.isSuperHR:
            self.responsible_box.configure(state="disabled")
        self.request_info_text.set('Номер заявки: ' + internalID + '\n' +
                                   'Офис : ' + officeName + '\n' +
                                   'Подразделение: ' + departmentName
                                   )
        if statusID in (1,4):
            self.bt_upload.config(state=tk.DISABLED)
            self.status_box.configure(state="disabled")
            self.upload_btn_text.set("Выбрать файл")
        elif statusID == 2:
            self.bt_upload.config(state=tk.NORMAL)
            self.status_box.config(state=tk.DISABLED)
            self.upload_btn_text.set("Выбрать файл")
        elif statusID == 3:
            self.responsible_box.configure(state="disabled")
            self.bt_upload.config(state=tk.DISABLED)
            self.upload_btn_text.set("Файл добавлен")
            self.status_box.configure(state="readonly")
        self.status_box.set('Не выбрано')




    def _convert_date(self, date, output=None):
        """ Take date and convert it into output format.
            If output is None datetime object is returned.

            date: str in format '%d[./]%m[./]%y' or '%d[./]%m[./]%Y'.
            output: str or None, output format.
        """
        date = date.replace('/', '.')
        try:
            dat = datetime.strptime(date, '%d.%m.%y')
        except ValueError:
            dat = datetime.strptime(date, '%d.%m.%Y')
        if output:
            return dat.strftime(output)
        return dat

    def _update_request(self):
        messagetitle = 'Изменение заявки'
        is_validated = self._validate_request_creation(messagetitle)
        if not is_validated:
            return

        update_vacancy = {'id': self.request_id,
                          'modifiedUserID': self.UserID,
                          'responsibleID': self.responsibleID[self.responsible_box.current()],
                          'fileCV': self.upload_filename,
                          'statusID': None or self.customStatusID[self.status_box.current()]
                          }
        update_success = self.conn.update_request(**update_vacancy)
        if update_success == 1:
            messagebox.showinfo(
                messagetitle, 'Заявка обновлена'
            )
            self._clear()
            # Если заявку возвращают в работу - удаляем согласованное ранее резюме
            if self.customStatusID[self.status_box.current()] == 2:
                self._remove_uploaded_file()
            self.controller._show_frame('PreviewForm')
        else:
            # self._remove_upload_file()
            messagebox.showerror(
                messagetitle, 'Произошла ошибка при обновлении договора'
            )
            # МВЗ, Договор, Арендодатель, ЕГРПОУ, Описание

    def _convert_str_date(self, date):
        """ Take str and convert it into date format.
            date: str in format '%d[./]%m[./]%y' or '%d[./]%m[./]%Y'.
        """
        date_time_str = date
        date_time_obj = dt.datetime.strptime(date_time_str, '%Y-%m-%d')
        return date_time_obj.date()

    def _restraint_by_office(self, event):
        """ Shows mvz_sap that corresponds to chosen MVZ and restraint offices.
            If 1 office is available, choose it, otherwise make box active.
        """
        # tcl language has no notion of None or a null value, so use '' instead
        self.mvz_sap = self.get_mvzSAP(self.mvz_current.get()) or ''

    def _row1_pack(self):
        self.request_info_label.pack(side=tk.LEFT, anchor=tk.W)

    def _row2_pack(self):
        self.separator.pack(fill='x')

    def _row3_pack(self):
        self.responsible_label.pack(side=tk.LEFT)
        self.responsible_box.pack(side=tk.RIGHT, padx=17)

    def _row4_pack(self):
        self.attach_label.pack(side=tk.LEFT, padx=0)

    def _row5_pack(self):
        self.status_label.pack(side=tk.LEFT)
        self.status_box.pack(side=tk.RIGHT, padx=17)

    def _row6_pack(self):
        pass


    def _top_pack(self):
        self.main_label.pack(side=tk.TOP, expand=False, anchor=tk.NW)

    def _validate_request_creation(self, messagetitle):
        """ Check if all fields are filled properly. """
        if not self.responsible_box.get():
            messagebox.showerror(
                messagetitle, 'Не выбран тип бизнеса'
            )
            return False
        return True


class PreviewForm(RecruitingFrame):
    def __init__(self, parent, controller, connection, user_info,
                 office, responsible, status_list, **kwargs):
        super().__init__(parent, controller, connection, user_info, office)

        self.statusID, self.status_list = zip(*[(None, 'Все'), ] + status_list)
        self.responsibleID, self.responsible = zip(
            *[(None, 'Все'), ] + responsible)
        self.UserID = self.user_info.UserID
        self.userOfficeID = self.user_info.officeID
        self.userDepartmentID = self.user_info.departmentID
        self.isHR = self.user_info.isHR

        # List of functions to get payments
        # determines what payments will be shown when refreshing
        self.vacancies_list = [self._get_all_vacancies]
        self.get_vacancies = self._get_all_vacancies
        # Parameters for sorting
        self.rows = None  # store all rows for sorting and redrawing
        self.sort_reversed_index = None  # reverse sorting for the last sorted column
        self.month = list(month_name)
        self.month_default = self.month[datetime.now().month]

        # Top Frame with description and user name
        top = tk.Frame(self, name='top_cf', padx=5)

        main_label = tk.Label(top, text='Просмотр заявок на персонал',
                              padx=5, font=('Calibri', 10, 'bold'))
        main_label.pack(side=tk.LEFT, expand=False, anchor=tk.NW)

        self._add_copyright(top)
        self._add_user_label(top)

        top.pack(side=tk.TOP, fill=tk.X, expand=False)

        # Filters
        filterframe = ttk.LabelFrame(self, text=' Фильтры ', name='filterframe')

        # First Filter Frame with (MVZ, office)
        row1_cf = tk.Frame(filterframe, name='row1_cf', padx=15, pady=5)

        self.office_label = tk.Label(row1_cf, text='Офис', padx=10)
        self.office_box = ttk.Combobox(row1_cf, width=40, state='readonly')
        self.office_box['values'] = list(self.office)

        self.responsible_label = tk.Label(row1_cf, text='Ответственный от HR',
                                          padx=20)
        self.responsible_box = ttk.Combobox(row1_cf, width=15,
                                            state='readonly')
        self.responsible_box['values'] = self.responsible

        self.status_label = tk.Label(row1_cf, text='Статус заявки', padx=20)
        self.status_box = ttk.Combobox(row1_cf, width=15, state='readonly')
        self.status_box['values'] = self.status_list

        self.bt3_1 = ttk.Button(row1_cf, text="Применить фильтр", width=20,
                                command=self._use_filter_and_refresh)
        self.bt3_2 = ttk.Button(row1_cf, text="Очистить фильтр", width=20,
                                command=self._clear_filters)

        # Pack row1_cf
        self._row1_pack()
        row1_cf.pack(side=tk.TOP, fill=tk.X)

        # Second Fill Frame with (Plan date, Sum, Tax)
        row2_cf = tk.Frame(filterframe, name='row2_cf', padx=15)

        # Pack row2_cf
        self._row2_pack()
        row2_cf.pack(side=tk.TOP, fill=tk.X)
        filterframe.pack(side=tk.TOP, fill=tk.X, expand=False, padx=10, pady=5)

        # Third Fill Frame (checkbox + button to apply filter)
        row3_cf = tk.Frame(filterframe, name='row3_cf', padx=15)

        # Pack row3_cf
        self._row3_pack()
        row3_cf.pack(side=tk.TOP, fill=tk.X, pady=10)

        # Set all filters to default
        self._clear_filters()

        # Text Frame
        preview_cf = ttk.LabelFrame(self, text=' Список заявок ',
                                    name='preview_cf')

        # column name and width
        self.headings = {'№ п/п': 40, 'ID': 0, 'Номер заявки': 90, 'UserID': 0,
                         'Офис': 250,
                         'Департамент': 0,
                         'Инициатор': 120, 'Дата внесения': 90,
                         'Плановая дата': 90,
                         'Должность кандидата': 0, 'ResponsibleUserID': 0,
                         'Ответственный от HR': 120,
                         'StatusID': 0, 'Тип заявки': 50, 'Статус': 80,
                         'Комментарий': 0,
                         'Файл заявки': 0, 'Файл резюме': 0
                         }

        self.table = ttk.Treeview(preview_cf, show='headings',
                                  selectmode='browse',
                                  style='HeaderStyle.Treeview'
                                  )

        self._init_table(preview_cf)
        self.table.pack(expand=tk.YES, fill=tk.BOTH)

        # asserts for headings used through script as indices
        head = self.table["columns"]
        msg = 'Heading order must be reviewed. Wrong heading: '
        assert head[1] == 'ID', '{}ID'.format(msg)
        assert head[-4] == 'Статус', '{}Статус'.format(msg)

        # Bottom Frame with buttons
        bottom_cf = tk.Frame(self, name='bottom_cf')
        # Show create buttons only for users with rights
        if self.user_info.isAccess in (1, 2):
            bt1 = ttk.Button(bottom_cf, text="Создать заявку", width=25,
                             command=lambda: controller._show_frame(
                                 'CreateForm'))
            bt1.pack(side=tk.LEFT, padx=10, pady=10)

            if self.isHR or self.UserID == 1:
                bt3 = ttk.Button(bottom_cf, text="Управление заявкой",
                                 width=30,
                                 command=self._edit_current_request)
                bt3.pack(side=tk.LEFT, padx=10, pady=10)

        bt6 = ttk.Button(bottom_cf, text="Выход", width=10,
                         command=controller._quit)
        bt6.pack(side=tk.RIGHT, padx=10, pady=10)

        bt5 = ttk.Button(bottom_cf, text="Подробно", width=10,
                         command=self._show_detail)
        bt5.pack(side=tk.RIGHT, padx=10, pady=10)

        bt4 = ttk.Button(bottom_cf, text="Экспорт в Excel", width=15,
                         command=self._export_to_excel)
        bt4.pack(side=tk.RIGHT, padx=10, pady=10)

        # Pack frames
        bottom_cf.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        preview_cf.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)

    def _add_copyright(self, parent):
        """ Adds user name in the top right corner. """
        copyright_label = tk.Label(parent, text="О программе",
                                   font=('Arial', 8, 'underline'),
                                   cursor="hand2")
        copyright_label.bind("<Button-1>", self._show_about)
        copyright_label.pack(side=tk.RIGHT, anchor=tk.N)

    def _center_popup_window(self, newlevel, w, h, static_geometry=True):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        start_x = int((screen_width / 2) - (w / 2))
        start_y = int((screen_height / 2) - (h * 0.7))

        if static_geometry == True:
            newlevel.geometry('{}x{}+{}+{}'.format(w, h, start_x, start_y))
        else:
            newlevel.geometry('+{}+{}'.format(start_x, start_y))

    def _change_preview_state(self):
        """ Change payments state that determines which payments will be shown.
        """
        self.get_vacancies = self._get_all_vacancies

    def _clear_filters(self):
        # self.initiator_box.set('Все')
        self.office_box.set('Все')
        self.responsible_box.set('Все')
        self.status_box.set('Все')

    def _edit_current_request(self):
        """ Raises UpdateForm with partially filled labels/entries. """
        curRow = self.table.focus()

        if curRow:
            # extract info to be putted in CreateForm
            to_fill = dict(zip(self.table["columns"],
                               self.table.item(curRow).get('values')))
            request_status = to_fill['Статус']
            if request_status == 'Отменена':
                messagebox.showinfo(
                    'Изменение заявки',
                    'Завка отменена и изменение в ней данных невозможно.'
                )
                return
            if request_status == 'Верифицировано':
                messagebox.showinfo(
                    'Изменение заявки',
                    'Завка уже верифицирована и изменение в ней данных невозможно.'
                )
                return
            self.controller._fill_UpdateForm(**to_fill)
            self.controller._show_frame('UpdateForm')

    def _delete_current(self):
        """ Raises CreateForm with partially filled labels/entries. """
        curRow = self.table.focus()

        if curRow:
            # extract info to be putted in CreateForm
            to_fill = dict(zip(self.table["columns"],
                               self.table.item(curRow).get('values')))

    def _export_to_excel(self):
        if not self.rows:
            return
        headings = {k: v for k, v in self.headings.items() if k != '№ п/п'}
        isExported = export_to_excel(headings, self.rows)
        if isExported:
            messagebox.showinfo(
                'Экспорт в Excel',
                'Данные экспортированы на рабочий стол'
            )
        else:
            messagebox.showerror(
                'Экспорт в Excel',
                'При экспорте произошла непредвиденная ошибка'
            )

    def _get_all_vacancies(self):
        """ Extract information from filters and get payments list. """
        filters = {
            'statusID': (self.statusID[self.status_box.current()]),
            'officeID': (self.officeID[self.office_box.current()]),
            'responsibleID': (
                self.responsibleID[self.responsible_box.current()]),
            'userOfficeID': self.userOfficeID,
            'userDepartmentID': self.userDepartmentID,
            'isHR': 1 if self.isHR else 0,
            'UserID': self.UserID
        }
        self.rows = self.conn.get_vacancies_list(**filters)

    def _init_table(self, parent):
        """ Creates treeview. """
        if isinstance(self.headings, dict):
            self.table["columns"] = tuple(self.headings.keys())
            self.table["displaycolumns"] = tuple(k for k in self.headings.keys()
                                                 if
                                                 k not in ('ID', 'Департамент',
                                                           'UserID',
                                                           'Должность кандидата',
                                                           'ResponsibleUserID',
                                                           'StatusID',
                                                           'Комментарий',
                                                           'Файл заявки',
                                                           'Файл резюме'))

            for head, width in self.headings.items():
                self.table.heading(head, text=head, anchor=tk.CENTER)
                self.table.column(head, width=width, anchor=tk.CENTER)

        else:
            self.table["columns"] = self.headings
            self.table["displaycolumns"] = self.headings
            for head in self.headings:
                self.table.heading(head, text=head, anchor=tk.CENTER)
                self.table.column(head, width=50 * len(head), anchor=tk.CENTER)

        for tag, bg, color in zip(self.status_list[1:6], (
                '#FFFFCC', '#bbded6', '#ffb6b9', '#8DB87C', '#CCCCCC'), (
                '#000000', '#000000', '#000000','#000000', '#000000')):
            self.table.tag_configure(tag, background=bg, foreground=color)

        self.table.bind('<Double-1>', self._show_detail)
        self.table.bind('<Button-1>', self._sort, True)

        scrolltable = tk.Scrollbar(parent, command=self.table.yview)
        self.table.configure(yscrollcommand=scrolltable.set)
        scrolltable.pack(side=tk.RIGHT, fill=tk.Y)


    def _raise_Toplevel(self, frame, title, width, height,
                        static_geometry=True, options=()):
        """ Create and raise new frame with limits.
        Input:
        frame - class, Frame class to be drawn in Toplevel;
        title - str, window title;
        width - int, width parameter to center window;
        height - int, height parameter to center window;
        static_geometry - bool, if True - width and height will determine size
            of window, otherwise size will be determined dynamically;
        options - tuple, arguments that will be sent to frame.
        """
        newlevel = tk.Toplevel(self.parent)
        # newlevel.transient(self)  # disable minimize/maximize buttons
        newlevel.title(title)
        newlevel.bind('<Escape>', lambda e, w=newlevel: w.destroy())
        frame(newlevel, *options)
        newlevel.resizable(width=False, height=False)
        self._center_popup_window(newlevel, width, height, static_geometry)
        newlevel.focus()
        newlevel.grab_set()

    def _refresh(self):
        """ Refresh information about vacancies. """
        try:
            self.get_vacancies()
        except MonthFilterError as e:
            messagebox.showerror(self.controller.title(), e.message)
            return
        self._show_rows(self.rows)

    def _resize_columns(self):
        """ Resize columns in treeview. """
        self.table.column('#0', width=36)
        for head, width in self.headings.items():
            self.table.column(head, width=width)

    def _row1_pack(self):
        self.office_label.pack(side=tk.LEFT)
        self.office_box.pack(side=tk.LEFT, padx=5, pady=5)
        self.responsible_label.pack(side=tk.LEFT)
        self.responsible_box.pack(side=tk.LEFT, padx=5, pady=5)
        self.status_label.pack(side=tk.LEFT)
        self.status_box.pack(side=tk.LEFT, padx=5, pady=5)

        self.bt3_2.pack(side=tk.RIGHT, padx=10)
        self.bt3_1.pack(side=tk.RIGHT, padx=10)

    def _row2_pack(self):
        pass

    def _row3_pack(self):
        pass

    def _show_about(self, event=None):
        """ Raise frame with info about app. """
        self._raise_Toplevel(frame=AboutFrame,
                             title='Заявки на поиск персонала ',
                             width=400, height=150)

    def _show_detail(self, event=None):
        """ Show details when double-clicked on row. """
        show_detail = (not event or (self.table.identify_row(event.y) and
                                     int(self.table.identify_column(event.x)[
                                         1:]) > 0
                                     )
                       )
        if show_detail:
            curRow = self.table.focus()
            if curRow:
                newlevel = tk.Toplevel(self.parent)
                newlevel.withdraw()
                # newlevel.transient(self.parent)  # disable minimize/maximize buttons
                newlevel.title('Просмотр заявки на персонал')
                newlevel.iconbitmap('resources/file.ico')
                newlevel.bind('<Escape>', lambda e, w=newlevel: w.destroy())
                DetailedPreview(newlevel, self, self.conn, self.userID,
                                self.headings,
                                self.table.item(curRow).get('values'),
                                self.table.item(curRow).get('tags'))
                newlevel.resizable(width=False, height=False)
                # width is set implicitly in DetailedPreview._newRow
                # based on columnWidths values
                self._center_popup_window(newlevel, 700, 300,
                                          static_geometry=False)
                newlevel.deiconify()
                newlevel.focus()
                newlevel.grab_set()
        else:
            # if double click on header - redirect to sorting rows
            self._sort(event)

    def _sort(self, event):
        if self.table.identify_region(event.x,
                                      event.y) == 'heading' and self.rows:
            # determine index of displayed column
            disp_col = int(self.table.identify_column(event.x)[1:]) - 1
            if disp_col < 1:  # ignore sort by '№ п/п' and checkboxes
                return
            # determine index of this column in self.rows
            # substract 1 because of added '№ п/п' which don't exist in data
            sort_col = self.table["columns"].index(
                self.table["displaycolumns"][disp_col]) - 1
            self.rows.sort(key=lambda x: x[sort_col],
                           reverse=self.sort_reversed_index == sort_col)
            # store index of last sorted column if sort wasn't reversed
            self.sort_reversed_index = None if self.sort_reversed_index == sort_col else sort_col
            self._show_rows(self.rows)

    def _show_rows(self, rows):
        """ Refresh table with new rows. """
        self.table.delete(*self.table.get_children())

        if not rows:
            return
        for i, row in enumerate(rows):
            self.table.insert('', tk.END,
                              values=(i + 1,) + tuple(
                                  map(lambda val: self._format_float(val)
                                  if isinstance(val, Decimal) else val, row)),
                              tags=(row[-4], 'unchecked'))

    def _use_filter_and_refresh(self):
        """ Change state to filter usage. """
        self._change_preview_state()
        self._refresh()


class DetailedPreview(tk.Frame):
    """ Class that creates Frame with information about chosen request. """

    def __init__(self, parent, parentform, conn, userID, head, info, tags):
        super().__init__(parent)
        self.parent = parent
        self.parentform = parentform
        self.conn = conn
        self.userID = userID
        self.rowtags = tags
        self.initiatorID = info[3]
        self.ID = info[1]
        self.statusID = info[12]
        self.filename_preview = str()
        self.cv_preview = str()
        # Top Frame with description and user name
        self.top = tk.Frame(self, name='top_cf', padx=5, pady=5)

        # Create a frame on the canvas to contain the buttons.
        self.table_frame = tk.Frame(self.top)

        # Add info to table_frame
        fonts = (('Arial', 9, 'bold'), ('Arial', 10))
        for row in zip(range(len(head)), zip(head, info)):
            if row[1][0] not in ('№ п/п', 'UserID', 'ID',
                                 'ResponsibleUserID', 'StatusID'):
                if row[1][0] == 'Файл заявки' and (row[1][1] != '-'
                                                   or row[1][1] is not None):
                    self.filename_preview = row[1][1]
                if row[1][0] == 'Файл резюме' and (row[1][1] != '-'
                                                   or row[1][1] is not None):
                    self.cv_preview = row[1][1]
                self._newRow(self.table_frame, fonts, *row)

        # self.appr_label = tk.Label(self.top, text='Адреса по договору',
        #                            padx=10, pady=5, font=('Arial', 10, 'bold'))

        # Top Frame with list mvz
        # self.appr_cf = tk.Frame(self, name='appr_cf', padx=5)
        #
        # # Add list of all mvz for current contract
        # fonts = (('Arial', 10), ('Arial', 10))
        # approvals = self.conn.get_additional_objects(self.contractID)
        # for rowNumber, row in enumerate(approvals):
        #     self._newRow(self.appr_cf, fonts, rowNumber + 1, row)

        self._add_buttons()
        self._pack_frames()

    def _open_file(self):
        pathToFile = UPLOAD_PATH + "\\" + self.filename_preview
        return os.startfile(pathToFile)

    def _open_cv(self):
        pathToFile = UPLOAD_PATH + "\\" + self.cv_preview
        return os.startfile(pathToFile)

    def _add_buttons(self):
        # Bottom Frame with buttons
        self.bottom = tk.Frame(self, name='bottom')
        if self.filename_preview:
            bt1 = ttk.Button(self.bottom, text="Требования", width=15,
                             command=self._open_file)
            bt1.pack(side=tk.LEFT, padx=5, pady=5)
        if self.cv_preview and self.cv_preview != '-':
            bt4 = ttk.Button(self.bottom, text="Резюме", width=15,
                             command=self._open_cv)
            bt4.pack(side=tk.LEFT, padx=5, pady=5)
        bt2 = ttk.Button(self.bottom, text="Закрыть", width=10,
                         command=self.parent.destroy)
        bt2.pack(side=tk.RIGHT, padx=5, pady=0)

        # show cancel button for initiator users
        if self.userID == self.initiatorID and self.statusID not in (4,5):
            bt3 = ttk.Button(self.bottom, text="Отменить заявку", width=18,
                             command=self._delete, style='ButtonRed.TButton')
            bt3.pack(side=tk.RIGHT, padx=5, pady=0)

    def _delete(self):
        mboxname = 'Отмена заявки на поиск персонала'
        confirmed = messagebox.askyesno(title=mboxname,
                                        message='Вы уверены, что хотите отменить эту заявку?')
        if confirmed:
            self.conn.update_request(self.ID, self.userID, None, None, 5)
            messagebox.showinfo(mboxname, 'Заявка отменена')
            self.parentform._refresh()
            self.parent.destroy()

    def _newRow(self, frame, fonts, rowNumber, info):
        """ Adds a new line to the table. """

        numberOfLines = []  # List to store number of lines needed
        columnWidths = [23, 60]  # Width of the different columns in the table

        # Find the length and the number of lines of each element and column
        for index, item in enumerate(info):
            # minimum 1 line + number of new lines + lines that too long
            numberOfLines.append(1 + str(item).count('\n') +
                                 sum(floor(len(s) / columnWidths[index]) for s
                                     in str(item).split('\n'))
                                 )

        # Find the maximum number of lines needed
        lineNumber = max(numberOfLines)

        # Define labels (columns) for row
        def form_column(rowNumber, lineNumber, col_num, cell, fonts):
            col = tk.Text(frame, bg='white', padx=10)
            col.insert(1.0, cell)
            col.grid(row=rowNumber, column=col_num + 1, sticky='news')
            col.configure(width=columnWidths[col_num],
                          height=min(9, lineNumber),
                          font=fonts[col_num], state="disabled")
            if lineNumber > 9 and col_num == 1:
                scrollbar = tk.Scrollbar(frame, command=col.yview)
                scrollbar.grid(row=rowNumber, column=col_num + 2, sticky='nsew')
                col['yscrollcommand'] = scrollbar.set

        for col_num, cell in enumerate(info):
            form_column(rowNumber, lineNumber, col_num, cell, fonts)

    def _pack_frames(self):
        self.top.pack(side=tk.TOP, fill=tk.X, expand=False)
        self.bottom.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        # self.appr_cf.pack(side=tk.TOP, fill=tk.X)
        self.table_frame.pack()
        # self.appr_label.pack(side=tk.LEFT, expand=False)
        self.pack()


class AboutFrame(tk.Frame):
    """ Creates a frame with copyright and info about app. """

    def __init__(self, parent):
        super().__init__(parent)

        self.top = ttk.LabelFrame(self, name='top_af')

        logo = tk.PhotoImage(file='resources/paper.png')
        self.logo_label = tk.Label(self.top, image=logo)
        self.logo_label.image = logo  # keep a reference to avoid gc!

        self.copyright_text = tk.Text(self.top, bg='#f1f1f1',
                                      font=('Arial', 8), relief=tk.FLAT)
        hyperlink = HyperlinkManager(self.copyright_text)

        def link_instruction():
            path = 'resources\\README.pdf'
            os.startfile(path)

        self.copyright_text.insert(tk.INSERT,
                                   'Платформа для учёта договоров аренды \n'
                                   '(версия ' + __version__ + ')\n')
        self.copyright_text.insert(tk.INSERT, "\n")

        def link_license():
            path = 'resources\\LICENSE.txt'
            os.startfile(path)

        self.copyright_text.insert(tk.INSERT,
                                   'Copyright © 2020 Департамент аналитики\n'
                                   'Офис контролинга логистики\n')
        self.copyright_text.insert(tk.INSERT, 'MIT License',
                                   hyperlink.add(link_license))

        self.bt = ttk.Button(self, text="Закрыть", width=10,
                             command=parent.destroy)
        self.pack_all()

    def pack_all(self):
        self.bt.pack(side=tk.BOTTOM, pady=5)
        self.top.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10)
        self.logo_label.pack(side=tk.LEFT, padx=10)
        self.copyright_text.pack(side=tk.LEFT, padx=10)
        self.pack(fill=tk.BOTH, expand=True)


if __name__ == '__main__':
    from db_connect import DBConnect
    from collections import namedtuple

    UserInfo = namedtuple('UserInfo', ['UserID', 'ShortUserName',
                                       'AccessType', 'isSuperUser', 'GroupID',
                                       'PayConditionsID'])

    with DBConnect(server='s-kv-center-s59', db='AnalyticReports') as sql:
        try:
            app = RecruitingApp(connection=sql,
                                user_info=UserInfo(24, 'TestName', 2, 1, 1, 2),
                                office=[('20511RC191', '20511RC191', 'Офис'),
                                        ('40900A2595', '40900A2595', 'Офис')],
                                status_list=[(1, 'На согл.'), (2, 'Отозв.')]
                                )
            app.mainloop()
        except Exception as e:
            print(e)
            raise
    input('Press Enter...')
