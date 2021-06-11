# -*- coding: utf-8 -*-
"""
Created on Fri Jul 12 15:03:05 2019

@author: v.shkaberda
"""
from _version import upd_path
from contextlib import suppress
from log_error import writelog
from singleinstance import Singleinstance
from splash_screen import SplashScreen
from time import sleep
from tkinter import Label, PhotoImage
from tkRecruiting import NetworkError, UnexpectedError
from shutil import copy2
import os, sys, zlib

SOURCE = zlib.decompress(upd_path).decode()
ALREADY_UPDATED = []


def update_files(main_path, path, directories, files):
    # specify path in current working directory
    relative_path = path.replace(main_path, '.')
    for file in files:
        if (relative_path, file) not in ALREADY_UPDATED:
            copy2(os.path.join(path, file), relative_path)
            ALREADY_UPDATED.append((relative_path, file))
    # create new directory if needed
    with suppress(FileExistsError):
        for directory in directories:
            os.mkdir(os.path.join(relative_path, directory))


def versioned(fname):
    """ Function to convert folder name into version ('1.2.13' -> (1, 2, 13)).
    """
    try:
        return tuple(map(int, fname.split('.')))
    except ValueError:
        return (0,)


def check_updates_and_run_app():
    # Extract names of all directories. Name of directory means version of app.
    app_versions = next(os.walk(SOURCE))[1]
    # Determine current version of application
    try:
        with open('recruiting.inf', 'r') as f:
            version_info = f.readline()
            version_info = versioned(version_info)
    except FileNotFoundError:
        from _version import version_info
    # Check all new versions and sort in descending order
    new_versions = sorted((x for x in app_versions if versioned(x) > version_info),
                          key=versioned,
                          reverse=True)
    for v in new_versions:
        path = os.path.join(SOURCE, v)
        for data in os.walk(path):
            update_files(path, *data)
    # Update version in recruiting.inf
    if new_versions:
        with open('recruiting.inf', 'w') as f:
            f.write(new_versions[0])
    # Run main executable
    os.startfile("recruiting.exe")
    sleep(5)


def main():
    # Create splash screen
    exception_handlers = {'NetworkError': NetworkError,
                          'UnexpectedError': UnexpectedError}
    root = SplashScreen(func=check_updates_and_run_app,
                        exception_handlers=exception_handlers)
    root.overrideredirect(True)

    logo = PhotoImage(file='resources/file-128.png')
    logo_label = Label(root, image=logo)
    logo_label.pack(side='top', pady=80)

    copyright_label = Label(root, text='© 2021 Офис прогнозирования \n Департамент мастер-данных и отчетности')
    copyright_label.pack(side='bottom', pady=15)

    label = Label(root,
                  text='Выполняется поиск обновлений и запуск приложения...')
    label.pack(expand='yes')

    root.after(300, root.task)
    root.mainloop()


if __name__ == '__main__':
    try:
        fname = os.path.basename(__file__)
        myapp = Singleinstance(fname)
        if myapp.aleradyrunning():
            sys.exit()
        main()
    except Exception as e:
        writelog(e)
        print(e)
    finally:
        sys.exit()