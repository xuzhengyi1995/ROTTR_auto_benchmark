import json
import os
import re
import time

import cv2
import imutils
import numpy as np
import win32com.client
import win32con

import settings
import win32api
import win32gui
import win32process
import win32ui

ROTTR_RES_FOLDER = os.path.join(
    os.environ['USERPROFILE'], 'Documents\\Rise of the Tomb Raider\\')
COMPUTER_NAME = os.environ['COMPUTERNAME']
RES = {}


def find_button(image):
    '''
    Use different button template to fit the image and caculate the coordinate of the button.
    '''
    templates = ['button.bmp', 'button_2.bmp', 'button_3.bmp']
    threshold = 0.6
    max_value = -1
    max_scale = 1
    max_location = None
    max_template = ''
    # Use different template
    for i in templates:
        template = cv2.imread(i, 0)
        w, h = template.shape[::-1]
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        # Also use different scale to test the template
        for scale in np.linspace(0.2, 1.8, 21)[::-1]:
            resized_image = imutils.resize(
                gray, width=int(gray.shape[1] * scale))
            if resized_image.shape[0] < w or resized_image.shape[1] < h:
                break
            result = cv2.matchTemplate(
                resized_image, template, cv2.TM_CCOEFF_NORMED)
            (_, local_max_val, _, loca_max_loc) = cv2.minMaxLoc(result)
            if local_max_val > threshold and local_max_val > max_value:
                max_value = local_max_val
                max_location = loca_max_loc
                max_scale = scale
                max_template = i
        if max_value >= 0.8:
            break

    if max_location:
        location = (max_location[0] + w / 2,
                    max_location[1] + h / 2) / max_scale
        location = (int(location[0]), int(location[1]))
        print('Found button \'%s\' at %f scale!' % (max_template, max_scale))
        print('Score:')
        print(max_value)
        print('Client location at:')
        print(location)
        return location
    else:
        print('Can not find button')
        return None


def click():
    '''
    Simulate mouse left button click event.
    '''
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.2)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)


def press_key(hdl, key):
    '''
    Simulate press key event.[Finally Not used in the script]
    '''
    win32api.SendMessage(hdl, win32con.WM_KEYDOWN, key, 0)
    time.sleep(0.2)
    win32api.SendMessage(hdl, win32con.WM_KEYUP, key, 0)


def close_window(hdl):
    '''
    Send close message to program.[Finall Not used in the script, the script kill the process directly.]
    '''
    win32api.SendMessage(hdl, win32con.WM_CLOSE, None, None)
    time.sleep(1)
    press_key(hdl, win32con.VK_RETURN)


def start_test(hdl):
    '''
    Start test, click the button.
    '''
    # Get the client rect.
    x_1, y_1, x_2, y_2 = win32gui.GetClientRect(hdl)
    if abs(x_2 / y_2 - 1.33) > abs(x_2 / y_2 - 1.77):
        print('16:9')
        x_f = 0.1951
        y_f = 0.515367
    else:
        print('4:3')
        x_f = 0.265495
        y_f = 0.506408
    # Caculate the button position
    bt_x = int(x_1 + (x_2 - x_1) * x_f)
    bt_y = int(y_1 + (y_2 - y_1) * y_f)
    print('Caculated position (%d,%d)' % (bt_x, bt_y))
    # Use opencv to find the button
    res = get_button_position(hdl)
    if res:
        bt_x, bt_y = res
    # Convert the client coordinate to screen coordinate.
    bt_x, bt_y = win32gui.ClientToScreen(hdl, (bt_x, bt_y))
    print('Button is at (%d,%d)' % (bt_x, bt_y))
    # Set the position
    win32api.SetCursorPos((bt_x, bt_y))
    time.sleep(2)
    print('Click button to start...')
    click()
    time.sleep(2)
    click()
    time.sleep(2)
    click()
    time.sleep(5)


def delete_old_result():
    '''
    Delete the old result files on the computer.
    '''
    files = os.listdir(ROTTR_RES_FOLDER)
    for i in files:
        os.remove(ROTTR_RES_FOLDER + i)


def get_result(filename):
    '''
    Read the result file and get the information.
    '''
    re_res = {
        'min_fps': re.compile(r'Min FPS: ([.0-9]+)'),
        'max_fps': re.compile(r'Max FPS: ([.0-9]+)'),
        'avg_fps': re.compile(r'Average FPS: ([.0-9]+)'),
        'num_frames': re.compile(r'Num Frames: ([0-9]+)')
    }
    with open(ROTTR_RES_FOLDER + filename, 'r') as f:
        data = f.read()
        res = {}
        for i in re_res:
            res[i] = re_res[i].findall(data)[0]
    return res


def check_is_res_ok(test_name):
    '''
    Check one test is finished or not.
    '''
    files = os.listdir(ROTTR_RES_FOLDER)
    for i in files:
        if str(i).count('%s_%s_2' % (test_name, COMPUTER_NAME)) != 0:
            RES[test_name] = get_result(i)
            print('%s TEST DONE! Result:' % test_name)
            print(RES[test_name])
            return True
    return False


def check_res():
    '''
    Check test function, waiting for test done.
    '''
    test_names = ['SpineOfTheMountain', 'ProphetsTomb', 'GeothermalValley']
    for i in test_names:
        while not check_is_res_ok(i):
            time.sleep(5)


def close_game(hdl):
    '''
    Kill the game process.
    '''
    print('Close game...')
    tid, pid = win32process.GetWindowThreadProcessId(hdl)
    os.system('taskkill /F /PID %d' % pid)


def write_final_result():
    '''
    Write the finall result to the json file.
    '''
    filename = 'res_' + str(int(time.time())) + '.json'
    with open(filename, 'w') as f:
        json.dump(RES, f)
    print('Test finished, result:')
    print(RES)
    print('Result can be found in %s' % filename)


def set_focus_window(hdl):
    '''
    Set the game window to focus and then we can do next jobs.
    '''
    tid, pid = win32process.GetWindowThreadProcessId(hdl)
    m_tid = win32api.GetCurrentThreadId()

    try:
        win32process.AttachThreadInput(m_tid, tid, -1)
        win32gui.BringWindowToTop(hdl)
        shell = win32com.client.Dispatch('WScript.Shell')
        shell.SendKeys('%')

        win32gui.SetForegroundWindow(hdl)
        win32gui.SetFocus(hdl)
    except Exception as e:
        print(e)
        print('Can not set focus')
    win32process.AttachThreadInput(m_tid, tid, 0)


def print_window(hdl):
    '''
    Get game image.
    '''
    # Get window height and width
    x_1, y_1, x_2, y_2 = win32gui.GetWindowRect(hdl)
    w_w = x_2 - x_1
    w_h = y_2 - y_1
    # Get client height and width
    _, _, c_w, c_h = win32gui.GetClientRect(hdl)
    # Caculate client coordinater
    x = int((w_w - c_w) / 2)
    y = w_h - c_h
    caption_h = win32api.GetSystemMetrics(win32con.SM_CYCAPTION) +\
        win32api.GetSystemMetrics(win32con.SM_CYBORDER) +\
        win32api.GetSystemMetrics(win32con.SM_CYFRAME) + 3
    if caption_h > y:
        print('Running in full screen mode.')
        x = 0
        y = 0
    else:
        print('Running in window mode.')
        y = caption_h

    hdl_dc = win32gui.GetWindowDC(hdl)
    mfc_dc = win32ui.CreateDCFromHandle(hdl_dc)
    save_dc = mfc_dc.CreateCompatibleDC()
    n_bmp = win32ui.CreateBitmap()
    n_bmp.CreateCompatibleBitmap(mfc_dc, c_w, c_h)
    save_dc.SelectObject(n_bmp)
    save_dc.BitBlt((0, 0), (c_w, c_h), mfc_dc, (x, y), win32con.SRCCOPY)
    img_arr = n_bmp.GetBitmapBits(True)
    img = np.fromstring(img_arr, dtype='uint8')
    img.shape = (c_h, c_w, 4)
    return cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)


def get_button_position(hdl):
    '''
    Get the image of the game and send it to the find_button function to search the button.
    '''
    res = None
    cnt = 0
    while not res and cnt < 5:
        game_image = print_window(hdl)
        res = find_button(game_image)
        cnt += 1
        if not res:
            print(
                'Tried %d times, maybe game not ready, will try again after 5s...' % cnt)
            time.sleep(5)
    return res


def main():
    '''
    Main function.
    '''
    steam_path = settings.STEAM_PATH
    # App id for ROTTR, Can be find here: https://steamdb.info/app/391220/
    rottr_app_id = settings.ROTTR_APP_ID
    rottr_title = settings.ROTTR_TITLE
    wait_for_lanuch = settings.WAIT_FOR_LANUCH
    delete_old_result()
    print('Start game...')
    os.system('%s -applaunch %s -nolauncher' %
              (steam_path, rottr_app_id))
    print('Wait %d second for launch...' % wait_for_lanuch)
    time.sleep(wait_for_lanuch)
    rottr_hdl = 0
    cnt = 0
    while rottr_hdl == 0 and cnt < 6:
        rottr_hdl = win32gui.FindWindow(None, rottr_title)
        cnt += 1
        if rottr_hdl == 0:
            time.sleep(5)
    if rottr_hdl == 0:
        raise Exception('Can not find ROTTR window.')
    set_focus_window(rottr_hdl)
    time.sleep(5)
    start_test(rottr_hdl)
    check_res()
    write_final_result()
    close_game(rottr_hdl)


if __name__ == '__main__':
    main()
