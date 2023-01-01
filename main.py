import logging
import logging.config
import re
import yaml

from pynput import keyboard
from pynput.keyboard import Key
from pywinauto.application import Application, findwindows
from pywinauto import Desktop
from win32clipboard import *
from win32con import *
from pywinauto.keyboard import send_keys

# from win32gui import GetForegroundWindow
# from win32process import GetWindowThreadProcessId
# from pywinauto.controls.uiawrapper import UIAWrapper
# from time import sleep


LOGGER_NAME = "pynput_test"
global logger
pressedKeys = set()

EXCLUDED_KEYS = {Key.f1, keyboard.Key.f2, keyboard.Key.f3, keyboard.Key.f4, keyboard.Key.f5, keyboard.Key.f6,
                 keyboard.Key.f7, keyboard.Key.f8, keyboard.Key.f9, keyboard.Key.f10, keyboard.Key.f11,
                 keyboard.Key.f12}

CTRL_C = "'\\x03'"  # {keyboard.Key.ctrl_l, keyboard.KeyCode.from_char('c')}
CTRL_V = "'\\x16'"  # {keyboard.Key.ctrl, keyboard.KeyCode.from_char('v')}
CTRL_SHIFT_S = {"'\\x13'", "Key.shift", "Key.ctrl_l"}


def copy_text_and_save_page(message):
    # send_keys("^c")

    fixed_message = fix_clipboard_text()

    # app = Application().start("C:\Program Files\Sublime Text 3\sublime_text.exe")
    # app = Application().connect(path="C:\Program Files\Sublime Text 3\sublime_text.exe")
    # app = Application().connect(process=2680)

    app_conn = Application()
    try:
        app = app_conn.connect(title_re=".*Sublime Text.*")
    except findwindows.WindowAmbiguousError:
        wins = findwindows.find_windows(title=".*Sublime Text.*")
        app = app_conn.connect(handle=wins[0])
    except Exception as err:
        logger.error(f"Exception occurred {err=}, {type(err)=}", exc_info=True)
        return -1

    dlg = app.top_window()
    # dlg = app.window(title_re=".*Sublime Text.*")
    # dlg.print_control_identifiers()

    dlg.type_keys(fixed_message + '\r\n', with_spaces=True)
    send_keys("{ENTER down}{ENTER up}")

    # chrome_dir = 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
    # start_args = ' --force-renderer-accessibility --start-maximized https://www.google.com/'
    # app = Application(backend="uia").start(chrome_dir+start_args)

    app_conn = Application()
    try:
        app = app_conn.connect(title_re=".*Chrome*")
    except findwindows.WindowAmbiguousError:
        wins = findwindows.find_windows(title=".*Chrome.*")
        app = app_conn.connect(handle=wins[0])
    except Exception as err:
        logger.error(f"Exception occurred {err=}, {type(err)=}", exc_info=True)
        return -1

    dialogs = app.windows()
    try:
        dlg = app.top_window()
    #        dlg.print_control_identifiers()
    except Exception as err:
        logger.error(f"Exception occurred {err=}, {type(err)=}", exc_info=True)

    chrome_window = Desktop(backend="uia").window(class_name_re='Chrome')
    #    chrome_window.print_control_identifiers()

    desktop = Desktop(backend="uia")
    uiawrapper = desktop.windows(title_re=".* Google Chrome .*", control_type="Pane")[0]
    # sleep(5)
    # window = GetForegroundWindow()
    # id, pid = GetWindowThreadProcessId(window)
    app = Application(backend="uia").connect(process=uiawrapper.element_info.process_id, time_out=10)
    dlg = app.top_window()
    title = "Save Page WE - normal page\nTiene acceso a este sitio web"
    dlg.child_window(title=title, control_type="MenuItem").select()

    # sleep(5)
    # chrome_window.print_control_identifiers()
    dlg_save = dlg.child_window(title="Guardar como", control_type="Window")
    dlg_save.wait('visible', timeout=30, retry_interval=1)
    combo_save = dlg_save.child_window(title="Nombre:", control_type="Edit")
    combo_save.type_keys(fixed_message, with_spaces=True)
    btn_save = dlg_save.child_window(title="Guardar", control_type="Button").click_input()

    pressedKeys.clear()

    return None


def evaluate_pressed_key(key, pressed):
    if key in EXCLUDED_KEYS:
        logger.info("Excluded key: {}".format(key))
        return None

    strKey = str(key)
    if pressed and not strKey in pressedKeys:
        pressedKeys.add(strKey)
        logger.debug('key {0} pressed'.format(strKey))

        if strKey in CTRL_SHIFT_S:
            if all(k in pressedKeys for k in CTRL_SHIFT_S):
                logger.info('--------------------------CTRL+SHFIT+S pressed!')
                copy_text_and_save_page(get_clipboard_text())

        if strKey == CTRL_C:
            logger.info('--------------------------CTRL+C pressed!')
        elif strKey == CTRL_V:
            logger.info('--------------------------CTRL+V pressed!')
    elif not pressed and strKey in pressedKeys:
        pressedKeys.remove(strKey)
        # TODO if a special key is released, remove all keys
        logger.info('{0} released'.format(key))

        if key == keyboard.Key.esc:
            # Stop listener
            return False

    #    logger.debug("The variable, key is of type: {}".format(type(key)))
    logger.info("pressedKeys({}): {}".format(pressed, pressedKeys))


def on_press_test(key):
    logger.debug('KEY', key)
    try:
        logger.debug('CHAR', key.char)  # letters, numbers et
    except ValueError:
        logger.debug('NAME', keyboard.KeyCode().from_char(key.name))


def on_press(key):
    evaluate_pressed_key(key, True)


def on_release(key):
    return evaluate_pressed_key(key, False)


def init_logger():
    with open('config.yml', 'r') as f:
        config = yaml.safe_load(f.read())
        logging.config.dictConfig(config)
    return logging.getLogger(LOGGER_NAME)


def init():
    # Collect events until released
    with keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()

    # ...or, in a non-blocking fashion:
    listener = keyboard.Listener(on_press=on_press, on_release=on_release)
    listener.start()


def fix_clipboard_text():
    OpenClipboard()
    text = None
    try:
        text = GetClipboardData(CF_UNICODETEXT)
        text = re.sub('[\\\/\:\*\?\"<>\|]', '', text)  # Remove characters not accepted in a windows file name
        text = re.sub('\r\n', '_', text)
        text = re.sub('\n', '_', text)
    except Exception as err:
        logger.error(f"Exception occurred {err=}, {type(err)=}", exc_info=True)
        return -1

    logger.debug("Result text: {}".format(text))
    SetClipboardText(text, CF_UNICODETEXT)
    #    text_test=GetClipboardData(CF_UNICODETEXT)
    #    logger.debug("Clipboard text: {}".format(text_test))
    CloseClipboard()
    return text


def get_clipboard_text():
    OpenClipboard()
    text = None
    try:
        text = GetClipboardData(CF_UNICODETEXT)
    except Exception as err:
        logger.error(f"Exception occurred {err=}, {type(err)=}", exc_info=True)
    CloseClipboard()
    return text


if __name__ == '__main__':
    logger = init_logger()
    init()
