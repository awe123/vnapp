import win32api
import win32gui
import win32ui
import win32con
import win32process
import commctrl
import ctypes
import time
import threading
import autopy
import json
import io

with io.open('resources\\vninfo.json', encoding='utf-8') as vnjsonfile:
    vnjsonstructure = json.load(vnjsonfile)

main_window_title = vnjsonstructure['main window title']
clock_popup_title = vnjsonstructure['clock popup title']
servers_popup_title = vnjsonstructure['servers popup title']
line1png = vnjsonstructure['1 line bitmap']
line2png = vnjsonstructure['2 line bitmap']
icon_notconnected1 = autopy.bitmap.Bitmap.open(vnjsonstructure['icon_nc1'])
icon_connected1 = autopy.bitmap.Bitmap.open(vnjsonstructure['icon_c1'])
icon_notconnected2 = autopy.bitmap.Bitmap.open(vnjsonstructure['icon_nc2'])
icon_connected2 = autopy.bitmap.Bitmap.open(vnjsonstructure['icon_c2'])
line1marker = vnjsonstructure['line1marker']
line2marker = vnjsonstructure['line2marker']
red_color = win32api.RGB(vnjsonstructure['red'][0], vnjsonstructure['red'][1], vnjsonstructure['red'][2])
black_color = win32api.RGB(vnjsonstructure['black'][0], vnjsonstructure['black'][1], vnjsonstructure['black'][2])
wleft = vnjsonstructure['app window rect'][0]
wtop = vnjsonstructure['app window rect'][1]
wright = vnjsonstructure['app window rect'][2]
wbottom = vnjsonstructure['app window rect'][3]
fsize = vnjsonstructure['font size']
fname = vnjsonstructure['font name']
fweight = vnjsonstructure['font weight']
qwidth, qheight = vnjsonstructure['clockpopup rect'][0], vnjsonstructure['clockpopup rect'][1]
bitmapSize_x = vnjsonstructure['toolbar width']
bitmapSize_y = vnjsonstructure['toolbar height']
tbimagepath = vnjsonstructure['toolbar images']
serverslist = []
# detect line 1/2
if autopy.bitmap.capture_screen().count_of_bitmap(autopy.bitmap.Bitmap.open(line1marker),
                                                  rect=((1, 83), (761, 174))) > 0:
    servers_row_dict = vnjsonstructure['1 линия']
    swidth, height = vnjsonstructure['serverspopup rect 1 line'][0], vnjsonstructure['serverspopup rect 1 line'][1]
    icon_connected = icon_connected1
    icon_notconnected = icon_notconnected1
    servers_picture = autopy.bitmap.Bitmap.open(line1png)
elif autopy.bitmap.capture_screen().count_of_bitmap(autopy.bitmap.Bitmap.open(line2marker),
                                                    rect=((1, 83), (634, 174))) > 0:
    servers_row_dict = vnjsonstructure['2 линия']
    swidth, height = vnjsonstructure['serverspopup rect 2 line'][0], vnjsonstructure['serverspopup rect 2 line'][1]
    icon_connected = icon_connected2
    icon_notconnected = icon_notconnected2
    servers_picture = autopy.bitmap.Bitmap.open(line2png)
else:
    raise Exception("Перекрыта область сканирования")
x_pos_list = []
stations_list = []
ip_list = []
btntxt_list = []
png_list = []
for i in range(len(servers_row_dict)):
    if i % 2 == 0:
        x_pos_list.append(servers_row_dict[i])
        stations_list.append(servers_row_dict[i + 1][0])
        ip_list.append(servers_row_dict[i + 1][1])
        btntxt_list.append(servers_row_dict[i + 1][2])
# number of servers in json to number of bitmaps, search range in xline.png
for j in range(len(servers_row_dict)):
    if j % 2 == 0 and servers_row_dict[j] != -1:
        png_list.append(servers_row_dict[j])
png_list_length = len(png_list)
# print(png_list)
servers_list_dict_from_name_bool = dict.fromkeys(png_list, True)
# print(servers_list_dict_from_name_bool)

checkbox_list_butid = []
rows = round(len(btntxt_list) / 5)
if round((len(btntxt_list) / 5), 1) - rows != 0:
    rows += 1
cols = 5
index_stations_list = 0
for i in range(rows):
    i += 1
    for j in range(cols):
        j += 1
        if index_stations_list < len(btntxt_list):
            button_ij = 30000 + i * 100 + j
            checkbox_list_butid.append(button_ij)
            index_stations_list += 1
# print(checkbox_list_butid)
checkbox_dict_from_wparam_xpos = dict(zip(checkbox_list_butid, x_pos_list))
# print(checkbox_dict_from_wparam_xpos)
app_window_refresher_thread_interval = vnjsonstructure["app_window_refresher_thread_interval"]
screen_bitmap_scanner_thread_interval = vnjsonstructure["screen_bitmap_scanner_thread_interval"]
sx, sy = vnjsonstructure["pos x serverslist wnd"], vnjsonstructure["pos y serverslist wnd"]
cmdX, cmdY = vnjsonstructure["pos x cmd launch wnd"], vnjsonstructure["pos y cmd launch wnd"]
cmd_wnd_color = vnjsonstructure["cmd window color"]
checkbox_btn_width, cmd_btn_width = vnjsonstructure["checkbox button width"], vnjsonstructure["cmd button width"]
buttons_height = vnjsonstructure["buttons height"]
buttons_margin_x, buttons_margin_y = vnjsonstructure["buttons margin x"], vnjsonstructure["buttons margin y"]
radio_btn_width, radio_btn_height = vnjsonstructure["radiobuttons width"], vnjsonstructure["radiobuttons height"]


class TBBUTTON(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('iBitmap', ctypes.c_int),
        ('idCommand', ctypes.c_int),
        ('fsState', ctypes.c_ubyte),
        ('fsStyle', ctypes.c_ubyte),
        ('bReserved', ctypes.c_ubyte * 2),
        ('dwData', ctypes.c_ulong),
        # c_size_t https://docs.python.org/3/library/ctypes.html#fundamental-data-types    https://www.viva64.com/ru/t/0062/
        ('iString', ctypes.c_int)
    ]


class ToolbarButton():
    """Toolbar buttons Clock, Cmd, Servers hide/show windows"""

    def __init__(self, title="", station="", ip=""):
        super(ToolbarButton, self).__init__()
        self.title = title
        self.station = station
        self.ip = ip

    def toggleshow(self, title):
        hwnd = win32ui.FindWindow(None, title)
        if hwnd.IsWindowVisible():
            hwnd.ShowWindow(win32con.SW_HIDE)
        else:
            hwnd.ShowWindow(win32con.SW_SHOW)

    def fastping(self):
        if displayinfo != "" and displayinfo != "OK":
            station = displayinfo.split("\n", 1)[0]
            ind = stations_list.index(station)
            ip = ip_list[ind]
            CreateCmdProcess().createwnd(station, ip)


class ClockPopup():
    """Toolbar button 1 > popup window Clock"""

    def __init__(self, title="", hWnd_q="", wndClassAtom="", qwidth="", qheight="", hWndApp="", hInstance=""):
        super(ClockPopup, self).__init__()
        self.title = title
        self.hWnd_q = hWnd_q
        self.wndClassAtom = wndClassAtom
        self.hWndApp = hWndApp
        self.hInstance = hInstance
        self.qwidth = qwidth
        self.qheight = qheight

    def app_terminate(self):
        ctypes.windll.user32.PostMessageA(win32gui.FindWindow(0, main_window_title), win32con.WM_CLOSE, 0, 0)

    def createwnd(self, hWndApp, qwidth, qheight, wndClassAtom, hInstance):
        apprect = qleft, qtop, qright, qbottom = win32gui.GetWindowRect(hWndApp)
        qx = qleft
        qy = qtop - qheight
        hWndclockpopup = win32gui.CreateWindowEx(
            win32con.WS_EX_PALETTEWINDOW,  # dwExStyle     WS_EX_OVERLAPPEDWINDOW WS_EX_TOOLWINDOW
            wndClassAtom,  # lpClassName
            clock_popup_title,  # lpWindowName
            win32con.WS_CAPTION | win32con.WS_SYSMENU,
            # WS_POPUP,         # dwStyle WS_POPUP  WS_OVERLAPPEDWINDOW  TBSTYLE_WRAPABLE !!! CCS_VERT TBSTYLE_FLAT = white
            qx,  # x
            qy,  # y
            qwidth,  # nWidth
            qheight,  # nHeight
            hWndApp,  # hWndParent
            0,  # hMenu
            hInstance,  # hInstance
            None  # lpParam
        )
        return hWndclockpopup

    def new_window_position_attached_to_parent(self, title):
        hwnd = win32ui.FindWindow(None, title)
        (wndleft, wndtop, wndright, wndbottom) = hwnd.GetClientRect()
        hWndApp = win32ui.FindWindow(None, main_window_title)
        (hleft, htop, hright, hbottom) = hWndApp.GetWindowRect()
        hwnd.MoveWindow((hleft, htop - wndbottom - 39, hleft + wndright + 16, htop), 1)

    def addradiobutton(self, title, hWnd_q, hInstance, row, col, butID):
        hWndradiobutton = win32gui.CreateWindow("BUTTON", title,
                                                win32con.WS_VISIBLE | win32con.WS_CHILD | win32con.BS_LEFT | win32con.BS_AUTORADIOBUTTON,
                                                buttons_margin_x * 2 + buttons_margin_x * 2 * col + radio_btn_width * (
                                                            col - 1),
                                                buttons_margin_y * 2 + buttons_margin_y * 2 * row + radio_btn_height * (
                                                            row - 1) + buttons_margin_y * 2,
                                                radio_btn_width,
                                                radio_btn_height,
                                                hWnd_q, butID, hInstance, None)
        return hWndradiobutton

    def addgroupbox(self, qwidth, hWnd_q, hInstance, butID):
        hWndgroupbox = win32gui.CreateWindow("BUTTON", None,
                                             win32con.WS_VISIBLE | win32con.WS_CHILD | win32con.BS_CENTER | win32con.BS_GROUPBOX | win32con.BS_BOTTOM,
                                             buttons_margin_x * 2, buttons_margin_y * 2, qwidth - buttons_margin_x * 8,
                                             buttons_margin_y * 8 + radio_btn_height * 2 + buttons_margin_y * 2, hWnd_q,
                                             20001, hInstance, None)
        return hWndgroupbox

    def addclosebutton(self, title, qwidth, qheight, hWnd_q, hInstance, butID):
        hWndclosebutton = win32gui.CreateWindow("BUTTON", "Закрыть",
                                                win32con.WS_VISIBLE | win32con.WS_CHILD | win32con.BS_CENTER | win32con.BS_DEFPUSHBUTTON,
                                                round(qwidth / 2) - round(radio_btn_width / 2),
                                                qheight - round(radio_btn_height / 2) - round((qheight - (
                                                            buttons_margin_y * 2 + radio_btn_height * 2 + buttons_margin_y * 2)) / 2),
                                                radio_btn_width, radio_btn_height, hWnd_q, 20000, hInstance, None)
        return hWndclosebutton


class CmdPopup():
    """Toolbar button 2 Cmd > launch CMD.exe PING failed server"""

    def __init__(self, wParam=""):
        super(CmdPopup, self).__init__()
        self.wParam = wParam

    def cmdlaunch(self, wParam):
        station_index = ((int(repr(wParam)[1:3])) - 1) * 5 + (int(repr(wParam)[3:5])) - 1
        station = stations_list[station_index]
        ip = ip_list[station_index]
        # print(((int(repr(wParam)[1:3]))-1)*5)
        # print(int(repr(wParam)[3:5]))
        # print(wParam, station_index)
        # print(wParam)
        CreateCmdProcess().createwnd(station, ip)


class CreateCmdProcess(object):
    """docstring for CreateCmdProcess"""

    def __init__(self, station="", ip=""):
        super(CreateCmdProcess, self).__init__()
        self.station = station
        self.ip = ip

    def createwnd(self, station, ip):
        cmdsettings = win32process.STARTUPINFO()
        cmdsettings.dwFlags = win32con.STARTF_USEPOSITION
        cmdsettings.dwX = cmdX
        cmdsettings.dwY = cmdY
        return win32process.CreateProcess(
            "C:\\Windows\\System32\\cmd.exe",
            "cmd.exe /k COLOR " + cmd_wnd_color + " && TITLE " + station + "&& PING " + ip,
            # 0A 1B https://ss64.com/nt/color.html
            None, None, False, win32con.CREATE_NEW_CONSOLE, None, None, cmdsettings)


class ServersPopup():
    """Toolbar button 3 Servers > list of servers with buttons"""

    def __init__(self, station="", swidth="", sheight="", hWndApp="", hWnd_s="", hInstance="", row="", col="", butID="",
                 wndClassAtom=""):
        super(ServersPopup, self).__init__()
        self.station = station
        self.swidth = swidth
        self.sheight = sheight
        self.hWndApp = hWndApp
        self.hWnd_s = hWnd_s
        self.hInstance = hInstance
        self.row = row
        self.col = col
        self.butID = butID
        self.wndClassAtom = wndClassAtom

    def createwnd(self, swidth, sheight, hWndApp, wndClassAtom, hInstance):
        apprect = sleft, stop, sright, sbottom = win32gui.GetWindowRect(hWndApp)
        # sx = round(sleft + (sright/2) - (swidth/2))
        # sy = stop - round(sheight/4)
        hWndserverspopup = win32gui.CreateWindowEx(
            win32con.WS_EX_PALETTEWINDOW,  # dwExStyle     WS_EX_OVERLAPPEDWINDOW WS_EX_TOOLWINDOW
            wndClassAtom,  # lpClassName
            servers_popup_title,  # lpWindowName
            win32con.WS_CAPTION | win32con.WS_SYSMENU,
            # dwStyle       WS_OVERLAPPEDWINDOW WS_POPUP     TBSTYLE_WRAPABLE !!! CCS_VERT TBSTYLE_FLAT = white
            sx,  # x
            sy,  # y
            swidth,  # nWidth
            sheight,  # nHeight
            hWndApp,  # hWndParent
            0,  # hMenu
            hInstance,  # hInstance
            None  # lpParam
        )
        return hWndserverspopup

    def addcheckboxcmdbuttons(self, station, hWnd_s, hInstance, row, col, butID):
        left = buttons_margin_x * col + (checkbox_btn_width + cmd_btn_width) * (
                    col - 1) + buttons_margin_x + buttons_margin_x
        top = buttons_margin_y * row + buttons_height * (
                    row - 1) + buttons_margin_y + buttons_margin_y + buttons_margin_y
        butIDcmd = butID + 10000  # Cmd butID= 40101, 40102, 40103 ... : 4[row:row][col:col]  "\n"+repr(butID)
        hWndcheckboxbutton = win32gui.CreateWindowEx(win32con.WS_EX_WINDOWEDGE, 'BUTTON', "",
                                                     win32con.WS_VISIBLE | win32con.WS_CHILD | win32con.BS_MULTILINE | win32con.BS_AUTOCHECKBOX | win32con.WS_EX_CLIENTEDGE,
                                                     # | win32con.WS_BORDER,
                                                     left, top, checkbox_btn_width, buttons_height, hWnd_s, butID,
                                                     hInstance, None)
        hwndcmdbutton = win32gui.CreateWindowEx(win32con.WS_EX_WINDOWEDGE, 'BUTTON', station,
                                                win32con.WS_VISIBLE | win32con.WS_CHILD | win32con.BS_BITMAP | win32con.BS_DEFPUSHBUTTON,
                                                # | win32con.WS_BORDER,
                                                left + checkbox_btn_width, top, cmd_btn_width, buttons_height, hWnd_s,
                                                butIDcmd, hInstance, None)
        # print(butID, butIDcmd)
        """uncomment to load button bitmap
        hBitmap = win32gui.LoadImage(0, "resources\\cmdbitmap.bmp", win32gui.IMAGE_BITMAP, 0, 0, 
            win32con.LR_LOADFROMFILE | win32con.LR_CREATEDIBSECTION)
        win32gui.SendMessage(hwndcmdbutton, 
            win32con.BM_SETIMAGE, win32gui.IMAGE_BITMAP, hBitmap)
        """
        return hWndcheckboxbutton, hwndcmdbutton

    def addgroupbox(self, swidth, sheight, hWndApp, hWnd_s, hInstance, butID):
        apprect = sleft, stop, sright, sbottom = win32gui.GetWindowRect(hWndApp)
        sx = round(sleft + (sright / 2) - (swidth / 2))
        sy = stop - round(sheight / 2)
        return win32gui.CreateWindow("BUTTON", None,
                                     win32con.WS_VISIBLE | win32con.WS_CHILD | win32con.BS_CENTER | win32con.BS_GROUPBOX | win32con.BS_BOTTOM,
                                     buttons_margin_x * 2, buttons_margin_y * 2, swidth - buttons_margin_x * 8,
                                     sheight - buttons_height - buttons_margin_y * 14, hWnd_s, butID, hInstance, None)

    def addclosebutton(self, swidth, sheight, hWndApp, station, hWnd_s, hInstance, butID):
        apprect = sleft, stop, sright, sbottom = win32gui.GetWindowRect(hWndApp)
        return win32gui.CreateWindow("BUTTON", "Закрыть",
                                     win32con.WS_VISIBLE | win32con.WS_CHILD | win32con.BS_CENTER | win32con.BS_DEFPUSHBUTTON,
                                     round(swidth / 2) - round((checkbox_btn_width + cmd_btn_width) / 2),
                                     sheight - buttons_height - buttons_margin_y * 10,
                                     checkbox_btn_width + cmd_btn_width, buttons_height, hWnd_s, 30000, hInstance, None)


class TextOnScreen():
    """srawing info"""

    def __init__(self):
        super(TextOnScreen, self).__init__()

    def scan(self):
        screen_picture = autopy.bitmap.capture_screen()
        notconnected_positions_tuple = screen_picture.find_every_bitmap(needle=icon_notconnected,
                                                                        tolerance=.1)  # , rect=((1, 83), (770, 174))
        # print("len =", len(notconnected_positions_tuple))
        # connected_positions_tuple = screen_picture.find_every_bitmap(icon_connected, rect=((1, 83), (634, 174)))
        # did not find bitmaps -> set "OK" black
        if len(notconnected_positions_tuple) == 0:
            displayinfo = "OK"
            txtcolor = black_color
        # if bitmaps have been found -> set multiline text, red
        else:
            displayinfo = ""
            txtcolor = red_color
            for crop_nc in range(len(notconnected_positions_tuple)):
                x1cn, y1cn, x2cn, y2cn = notconnected_positions_tuple[crop_nc][0] + 16, \
                                         notconnected_positions_tuple[crop_nc][1], 58, 16
                cropped_58x16_serv_notcon = screen_picture.cropped(rect=((x1cn, y1cn), (x2cn, y2cn)))
                for xpos in png_list:
                    # xpos = 74*crop_row+16
                    # xpos=crop_row
                    server_iterate_bitmap = servers_picture.cropped(rect=((xpos, 0), (58, 16)))
                    if cropped_58x16_serv_notcon.is_bitmap_equal(server_iterate_bitmap, 0.1) and \
                            servers_list_dict_from_name_bool[xpos]:
                        # print(servers_list_dict_from_name_bool[xpos])
                        displayinfo += servers_row_dict[servers_row_dict.index(xpos) + 1][0] + "\n"
            # print("displayinfo = ", displayinfo)
            if displayinfo == "":
                displayinfo = "OK"
                txtcolor = red_color
        return displayinfo, txtcolor


def main():
    """**********************************************************************
    *              MAIN WINDOW INIT
    **********************************************************************"""
    hInstance = win32gui.GetModuleHandle(None)
    wndClass = win32gui.WNDCLASS()
    wndClass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
    wndClass.lpfnWndProc = wndproc
    wndClass.hInstance = hInstance
    wndClass.hIcon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
    wndClass.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
    wndClass.hbrBackground = 5  # win32con.COLOR_BTNSHADOW # win32gui.GetStockObject(win32con.WHITE_BRUSH)
    wndClass.lpszClassName = 'pClassName'
    wndClassAtom = None
    wndClassAtom = win32gui.RegisterClass(wndClass)
    hWndApp = win32gui.CreateWindowEx(0, wndClassAtom, main_window_title, win32con.WS_OVERLAPPEDWINDOW,
                                      win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT,
                                      win32con.CW_USEDEFAULT, 0, 0, hInstance, None)
    win32gui.SetWindowPos(hWndApp, win32con.HWND_TOPMOST, wleft, wtop, wright, wbottom, win32con.SWP_SHOWWINDOW)
    win32gui.ShowWindow(hWndApp, win32con.SW_SHOWNORMAL)

    """**********************************************************************
    *              CLOCK WINDOW INIT - Toolbar button 2
    **********************************************************************"""
    # qwidth, qheight = 300, 200
    hWnd_q = ClockPopup().createwnd(hWndApp, qwidth, qheight, wndClassAtom, hInstance)
    hWndButton20101 = ClockPopup().addradiobutton("7:45 / 19:45", hWnd_q, hInstance, row=1, col=1, butID=20101)
    hWndButton20102 = ClockPopup().addradiobutton("7:45", hWnd_q, hInstance, row=1, col=2, butID=20102)
    hWndButton20201 = ClockPopup().addradiobutton("19:45", hWnd_q, hInstance, row=2, col=1, butID=20201)
    hWndButton20202 = ClockPopup().addradiobutton("Не выключать", hWnd_q, hInstance, row=2, col=2, butID=20202)
    hWndGrpBox20001 = ClockPopup().addgroupbox(qwidth, hWnd_q, hInstance, butID=20001)
    hWndButton20000 = ClockPopup().addclosebutton("Закрыть", qwidth, qheight, hWnd_q, hInstance, butID=20000)
    check_default_radiobutton_clockpopup = win32ui.FindWindow(None, clock_popup_title)
    check_default_radiobutton_clockpopup.CheckRadioButton(20101, 20202, 20101)

    """**********************************************************************
    *              SERVERS WINDOW INIT - Toolbar button 3
    **********************************************************************"""
    # swidth, height = 700, 500
    hWnd_s = ServersPopup().createwnd(swidth, height, hWndApp, wndClassAtom, hInstance)
    rows = round(len(btntxt_list) / 5)
    if round((len(btntxt_list) / 5), 1) - rows != 0:
        rows += 1
    cols = 5
    # print("rows, cols = ", rows, cols)
    index_stations_list = 0
    for i in range(rows):
        i += 1
        for j in range(cols):
            j += 1
            if index_stations_list < len(btntxt_list):
                # print(index_stations_list)
                button_ij = 30000 + i * 100 + j
                # print(btntxt_list[index_stations_list])
                # print(button_ij, btntxt_list[index_stations_list], index_stations_list , "i=", i, "j=",j)
                hWndcheckboxbutton, hwndcmdbutton = ServersPopup().addcheckboxcmdbuttons(
                    btntxt_list[index_stations_list], hWnd_s, hInstance, row=i, col=j, butID=button_ij)
                if btntxt_list[index_stations_list] == "":
                    win32gui.ShowWindow(hWndcheckboxbutton, win32con.SW_HIDE)
                    win32gui.ShowWindow(hwndcmdbutton, win32con.SW_HIDE)
                index_stations_list += 1
    hWndGrpBox30001 = ServersPopup().addgroupbox(swidth, height, hWndApp, hWnd_s, hInstance, butID=30001)
    hWndButton30000 = ServersPopup().addclosebutton(swidth, height, hWndApp, "Закрыть", hWnd_s, hInstance, butID=30000)

    """**********************************************************************
    *              TOOLBAR INIT
    **********************************************************************"""
    hImageList = None
    numButtons = 1
    # bitmapSize_x = 45
    # bitmapSize_y = 45
    # tbimagepath = {"Q":"resources/tbq.bmp", "C":"resources/tbc.bmp", "S":"resources/tbs.bmp"}

    hWndToolbar = win32gui.CreateWindowEx(
        0,  # dwExStyle     WS_EX_OVERLAPPEDWINDOW
        commctrl.TOOLBARCLASSNAME,  # lpClassName
        None,  # lpWindowName
        win32con.WS_CHILD | win32con.WS_VISIBLE | commctrl.TBSTYLE_WRAPABLE | commctrl.TBSTYLE_FLAT | commctrl.CCS_ADJUSTABLE,
        # dwStyle     TBSTYLE_WRAPABLE !!! CCS_VERT TBSTYLE_FLAT = white
        0,  # x
        0,  # y
        0,  # nWidth
        0,  # nHeight
        hWndApp,  # hWndParent
        10,  # hMenu
        hInstance,  # hInstance
        None  # lpParam
    )

    win32gui.SendMessage(hWndToolbar, commctrl.TB_BUTTONSTRUCTSIZE, ctypes.sizeof(TBBUTTON), 0)
    hImageList = win32gui.ImageList_Create(bitmapSize_x, bitmapSize_y, win32gui.ILC_COLOR24 | win32gui.ILC_MASK,
                                           numButtons, 0)
    win32gui.SendMessage(hWndToolbar, commctrl.TB_SETIMAGELIST, 0, hImageList)
    win32gui.SendMessage(hWndToolbar, commctrl.TB_SETBUTTONSIZE, 0, win32api.MAKELONG(bitmapSize_x, bitmapSize_y))
    tbstring = win32gui.SendMessage(hWndToolbar, commctrl.TB_ADDSTRING, hInstance, 33)
    tbButtons = TBBUTTON()
    tbButtons.bReserved = 0, 0
    tbButtons.dwData = 0
    tbButtons.iString = tbstring
    # Toolbar button 1
    hBitmap = win32gui.LoadImage(0, tbimagepath["clock"], win32gui.IMAGE_BITMAP, bitmapSize_x, bitmapSize_y,
                                 win32gui.LR_LOADFROMFILE)
    toolbarBitmap = win32gui.ImageList_Add(hImageList, hBitmap, 0)
    tbButtons.iBitmap = toolbarBitmap
    tbButtons.idCommand = 11
    tbButtons.fsState = commctrl.TBSTATE_ENABLED | commctrl.TBSTATE_WRAP
    tbButtons.fsStyle = commctrl.BTNS_BUTTON  # | commctrl.BTNS_CHECK
    win32gui.SendMessage(hWndToolbar, commctrl.TB_ADDBUTTONS, 1, tbButtons)
    # Toolbar button 2
    hBitmap = win32gui.LoadImage(0, tbimagepath["cmd"], win32gui.IMAGE_BITMAP, bitmapSize_x, bitmapSize_y,
                                 win32gui.LR_LOADFROMFILE)
    toolbarBitmap = win32gui.ImageList_Add(hImageList, hBitmap, 0)
    tbButtons.iBitmap = toolbarBitmap
    tbButtons.idCommand = 12
    tbButtons.fsState = commctrl.TBSTATE_ENABLED | commctrl.TBSTATE_WRAP
    tbButtons.fsStyle = commctrl.BTNS_BUTTON
    win32gui.SendMessage(hWndToolbar, commctrl.TB_ADDBUTTONS, 1, tbButtons)
    # Toolbar button 3
    hBitmap = win32gui.LoadImage(0, tbimagepath["servers"], win32gui.IMAGE_BITMAP, bitmapSize_x, bitmapSize_y,
                                 win32gui.LR_LOADFROMFILE)
    toolbarBitmap = win32gui.ImageList_Add(hImageList, hBitmap, 0)
    tbButtons.iBitmap = toolbarBitmap
    tbButtons.idCommand = 13
    tbButtons.fsState = commctrl.TBSTATE_ENABLED | commctrl.TBSTATE_WRAP
    tbButtons.fsStyle = commctrl.BTNS_BUTTON  # | commctrl.BTNS_WHOLEDROPDOWN
    win32gui.SendMessage(hWndToolbar, commctrl.TB_ADDBUTTONS, 1, tbButtons)
    win32gui.ShowWindow(hWndToolbar, win32con.SW_SHOWNORMAL)

    thread_scanner = threading.Thread(target=screen_bitmap_scanner_thread, args=(hWndApp,))
    thread_scanner.daemon = True
    thread_scanner.start()

    thread_refresher = threading.Thread(target=app_window_refresher_thread, args=(hWndApp,))
    thread_refresher.daemon = True
    thread_refresher.start()

    thread_clock = threading.Thread(target=app_terminator_scheduler_thread, args=(hWndApp,))
    thread_clock.daemon = True
    thread_clock.start()

    win32gui.PumpMessages()


def app_terminator_scheduler_thread(hWndApp):
    global schedule_time_clock_popup_radiobutton_id
    schedule_time_clock_popup_radiobutton_id = 20101
    while True:
        # tm_year=2020, tm_mon=6, tm_mday=20, tm_hour=19, tm_min=11, tm_sec=3, tm_wday=5, tm_yday=172, tm_isdst=0
        if schedule_time_clock_popup_radiobutton_id == 20101:  # 7:45 / 19:45
            if time.localtime().tm_hour == 19 or time.localtime().tm_hour == 7:
                if time.localtime().tm_min == 45:
                    ClockPopup().app_terminate()
        if schedule_time_clock_popup_radiobutton_id == 20102:  # 7:45
            if time.localtime().tm_hour == 7 and time.localtime().tm_min == 45:
                ClockPopup().app_terminate()
        if schedule_time_clock_popup_radiobutton_id == 20201:  # 19:45
            if time.localtime().tm_hour == 19 and time.localtime().tm_min == 45:
                ClockPopup().app_terminate()
        if schedule_time_clock_popup_radiobutton_id == 20202:  # Не выключать
            time.sleep(10)
        time.sleep(10)


def screen_bitmap_scanner_thread(hWndApp):
    global displayinfo, txtcolor, refreshercommander
    displayinfo = "OK"
    txtcolor = black_color
    refreshercommander = True
    while True:
        # send params to wndproc
        previous_displayinfo, previous_txtcolor = displayinfo, txtcolor
        displayinfo, txtcolor = TextOnScreen().scan()
        if displayinfo == previous_displayinfo and txtcolor == previous_txtcolor:
            refreshercommander = False
        else:
            refreshercommander = True
        # refresh interval
        # print(displayinfo, txtcolor)
        # print(refreshercommander)
        time.sleep(screen_bitmap_scanner_thread_interval)


def app_window_refresher_thread(hWndApp):
    while True:
        # refresh app main window
        time.sleep(app_window_refresher_thread_interval)
        if refreshercommander:
            # print("refreshing")
            (left, top, right, bottom) = win32gui.GetClientRect(hWndApp)
            win32gui.RedrawWindow(
                hWndApp,
                (0, 55, right, bottom),  # toolbar Y margin = 5 + toolbar + 5 = 55
                None,
                win32con.RDW_INVALIDATE | win32con.RDW_ERASE
            )


def wndproc(hWndApp, message, wParam, lParam):
    if message == win32con.WM_COMMAND:

        # TOOLBAR BUTTONS when pressed events
        # Toolbar button [1] CLOCK when pressed
        if wParam == 11:
            ClockPopup().new_window_position_attached_to_parent(title=clock_popup_title)
            ToolbarButton().toggleshow(title=clock_popup_title)
        # Toolbar button [2] CMD when pressed
        if wParam == 12:
            ToolbarButton().fastping()
        # Toolbar button [3] SERVERS when pressed
        if wParam == 13:
            ToolbarButton().toggleshow(title=servers_popup_title)

        # POPUP WINDOWS events
        # CLOCK [1]
        # button "Закрыть" > CLOCK
        if wParam == 20000:
            ToolbarButton().toggleshow(title=clock_popup_title)
        # button "выключать" > CLOCK
        # if wParam == 222:
        #     ctypes.windll.user32.PostMessageA(win32gui.FindWindow(0, main_window_title), win32con.WM_CLOSE, 0, 0)
        # radiobuttons
        if repr(wParam).startswith("2") and not repr(wParam).endswith("0"):
            hwndCheck = win32gui.GetDlgItem(hWndApp, wParam)
            radiobuttonStatus = win32gui.SendMessage(hwndCheck, win32con.BM_GETCHECK, 0, 0)
            global schedule_time_clock_popup_radiobutton_id
            if radiobuttonStatus == win32con.BST_CHECKED:
                schedule_time_clock_popup_radiobutton_id = wParam
            if radiobuttonStatus == win32con.BST_UNCHECKED:
                schedule_time_clock_popup_radiobutton_id = wParam

        # SERVERS [3]
        # button "Закрыть" > SERVERS
        if wParam == 30000:
            ToolbarButton().toggleshow(title=servers_popup_title)
        # buttons "cmd"
        if repr(wParam).startswith("4") and not repr(wParam).endswith("0"):
            CmdPopup().cmdlaunch(wParam)
        # checkboxes
        if repr(wParam).startswith("3") and not repr(wParam).endswith("0"):
            hwndCheck = win32gui.GetDlgItem(hWndApp, wParam)
            checkboxStatus = win32gui.SendMessage(hwndCheck, win32con.BM_GETCHECK, 0, 0)
            global servers_list_dict_from_name_bool
            if checkboxStatus == win32con.BST_CHECKED:
                # "\n"+repr(butID+10000)
                # print("xpos=", checkbox_dict_from_wparam_xpos.get(wParam))
                servers_list_dict_from_name_bool.pop(checkbox_dict_from_wparam_xpos.get(wParam))
                servers_list_dict_from_name_bool.update({checkbox_dict_from_wparam_xpos.get(wParam): False})
                # print(servers_list_dict_from_name_bool)
                # print(wParam, "checked") # 3rrcc
            if checkboxStatus == win32con.BST_UNCHECKED:
                servers_list_dict_from_name_bool.pop(checkbox_dict_from_wparam_xpos.get(wParam))
                servers_list_dict_from_name_bool.update({checkbox_dict_from_wparam_xpos.get(wParam): True})
                # print(wParam, "unchecked")

    # main window text drawing
    if message == win32con.WM_PAINT and not win32gui.IsWindow(
            win32gui.GetWindow(hWndApp, win32con.GW_OWNER)):  # hWndApp isolation
        hdc, paintStruct = win32gui.BeginPaint(hWndApp)
        # # print("b", paintStruct)
        # hdc, fErase, rcPaint, fRestore, fIncUpdate, rgbReserved = paintStruct
        # left, top, right, bottom = rcPaint
        # rcPaint = left, top+50, right, bottom-50
        # paintStruct = hdc, fErase, rcPaint, fRestore, fIncUpdate, rgbReserved
        # # print("a", paintStruct)
        dpiScale = win32ui.GetDeviceCaps(hdc, win32con.LOGPIXELSX) / 60.0
        fontSize = fsize
        lf = win32gui.LOGFONT()
        lf.lfFaceName = fname
        lf.lfHeight = int(round(dpiScale * fontSize))
        lf.lfWeight = fweight
        lf.lfQuality = win32con.DEFAULT_QUALITY
        hf = win32gui.CreateFontIndirect(lf)
        win32gui.SelectObject(hdc, hf)
        left, top, right, bottom = win32gui.GetClientRect(hWndApp)
        rect = left, top + 50, right, bottom
        win32gui.SetBkMode(hdc, win32con.TRANSPARENT)
        win32gui.SetTextColor(hdc, txtcolor)
        win32gui.DrawText(hdc, displayinfo, -1, rect, win32con.DT_CENTER | win32con.DT_VCENTER)
        win32gui.EndPaint(hWndApp, paintStruct)
        return 0
    # close all owned wiundows - popups
    if message == win32con.WM_CLOSE and win32gui.IsWindow(win32gui.GetWindow(hWndApp, win32con.GW_OWNER)):
        win32ui.FindWindow(None, clock_popup_title).ShowWindow(win32con.SW_HIDE)
    if message == win32con.WM_CLOSE and win32gui.IsWindow(win32gui.GetWindow(hWndApp, win32con.GW_OWNER)):
        win32ui.FindWindow(None, servers_popup_title).ShowWindow(win32con.SW_HIDE)
    # close app
    elif message == win32con.WM_DESTROY and win32gui.IsWindow(hWndApp):
        # print("WM_DESTROY", hWndApp)
        win32gui.PostQuitMessage(0)
        return 0
    else:
        return win32gui.DefWindowProc(hWndApp, message, wParam, lParam)


if __name__ == '__main__':
    main()