from _version import upd_path
from tkinter import Tk, messagebox
from win32com.client import Dispatch
import getpass, os, zipfile, zlib

SOURCE = zlib.decompress(upd_path).decode()
TARGET = os.path.join('C:\\Users', getpass.getuser(), 'AppData\\Local')
DESKTOP = os.path.join('C:\\Users', getpass.getuser(), 'Desktop')
WDIR = os.path.join(TARGET, 'Recruiting')
TARGETFILE = os.path.join(WDIR, 'recruiting_checker.exe')
ICONFILE = os.path.join(WDIR, 'resources\\file.ico')

print(TARGETFILE)

class SuccessMsg(Tk):
    """ Raise an error when user doesn't have permission to work with app.
    """
    def __init__(self):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showinfo(
            'Заявки на персонал',
            'Установка завершена.\n'
            'На рабочем столе создан ярлык для запуска.'
        )
        self.destroy()

def create_shortcut(path, target='', wDir='', icon=''):
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = wDir
    shortcut.Description = 'Заявки на персонал'
    if icon:
        shortcut.IconLocation = icon
    shortcut.save()

def main():
    print('Выполняется начальная установка и создание ярлыков...')
    # extract actual version of app
    with zipfile.ZipFile(os.path.join(SOURCE, 'Recruiting.zip'), 'r') as zip_ref:
        zip_ref.extractall(TARGET)
    create_shortcut(os.path.join(DESKTOP, 'Заявки на персонал.lnk'),
                    TARGETFILE,
                    WDIR,
                    ICONFILE)
    SuccessMsg()

if __name__ == '__main__':
    main()