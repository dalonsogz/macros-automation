from pywinauto.application import Application
from win32clipboard import *
from win32con import *
from time import sleep


def notepadTest():
    app = Application(backend="uia").start('notepad.exe')
    # describe the window inside Notepad.exe process
    dlg_spec = app.window(best_match='.*: Bloc de notas$')
    # wait till the window is really open
    actionable_dlg = dlg_spec.wait('visible')

    text_message = "Hi Aliens! Text coming from pywinauto!"
    actionable_dlg.type_keys(text_message, with_spaces=True)
    return app,actionable_dlg


def clipboardTest(dlg):
    val = "Clipboard text"
    OpenClipboard()
    SetClipboardText(val)
    SetClipboardText(val,CF_UNICODETEXT)
    text=GetClipboardData(CF_UNICODETEXT)
    CloseClipboard()
    dlg.type_keys("\r\n"+text, with_spaces=True)


def url2filename(url):
    url = url.encode('UTF-8')
    return base64.urlsafe_b64encode(url).decode('UTF-8')

def filename2url(f):
    return base64.urlsafe_b64decode(f).decode('UTF-8')


if __name__ == '__main__':
    app,dlg = notepadTest()
    clipboardTest(dlg)
    dlg.type_keys("\r\n\r\nClosing app...", with_spaces=True)
    sleep(4)
    app.kill()
