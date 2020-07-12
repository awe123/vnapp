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

# windows titles
main_window_title = vnjsonstructure['main window title']
clock_popup_title = vnjsonstructure['clock popup title']
servers_popup_title = vnjsonstructure['servers popup title']
# png file of 70 bitmaps in one row
# the file is used to find a bitmap match and get x position of the bitmap in the image  
line1_allservers_png = vnjsonstructure['1 line bitmap']
line2_allservers_png = vnjsonstructure['2 line bitmap']
# different bitmaps icon_notconnected for 1 and 2 lines because colors are not the same
# icon_connected images are not used, they're for testing purpose
icon_notconnected1 = autopy.bitmap.Bitmap.open(vnjsonstructure['icon_nc1'])
icon_connected1 = autopy.bitmap.Bitmap.open(vnjsonstructure['icon_c1'])
icon_notconnected2 = autopy.bitmap.Bitmap.open(vnjsonstructure['icon_nc2'])
icon_connected2 = autopy.bitmap.Bitmap.open(vnjsonstructure['icon_c2'])
# here is the detection of which line to process
# the process is based on distinguishing of bitmaps of first server in the VN GUI:
# 1st line - top left corner academ1.png, 2nd line - column 1 row 2 avtozav1.png
line1_detector_png = vnjsonstructure['line1marker']
line2_detector_png = vnjsonstructure['line2marker']
# colors red and black settings for text in main window
red_color = win32api.RGB(vnjsonstructure['red'][0], vnjsonstructure['red'][1], vnjsonstructure['red'][2])
black_color = win32api.RGB(vnjsonstructure['black'][0], vnjsonstructure['black'][1], vnjsonstructure['black'][2])
# main window rectangle, screen coords
appwnd_left = vnjsonstructure['app window rect'][0]
appwnd_top = vnjsonstructure['app window rect'][1]
appwnd_right = vnjsonstructure['app window rect'][2]
appwnd_bottom = vnjsonstructure['app window rect'][3]
# text settings appeared in the main window
font_size = vnjsonstructure['font size']
font_name = vnjsonstructure['font name']
font_weight = vnjsonstructure['font weight']
# clock popup width and height
clockwnd_width = vnjsonstructure['clockpopup rect'][0]
clockwnd_height = vnjsonstructure['clockpopup rect'][1]
# buttons width, height in main window - clock, cmd, servers
toolbar_button_size_x = vnjsonstructure['toolbar width']
toolbar_button_size_y = vnjsonstructure['toolbar height']
toolbar_image_path = vnjsonstructure['toolbar images']
# process of detection of line 1 or 2
# capture screen image and find in it a bitmap with academ1 or avtozav1 bitmap
# if bitmap is found, its' count is > 0
if autopy.bitmap.capture_screen().count_of_bitmap(autopy.bitmap.Bitmap.open(line1_detector_png),
                                                  rect=((1, 83), (761, 174))) > 0:
    # nested list with format: [int_xpos, [str_cmdtitle, str_ip, str_buttontext]]
    # x_pos - line1.png x coordinate of bitmap
    # str_cmdtitle - string value for the title in cmd popup
    # str_ip - value to insert in cmd string "ping 10.99.14.1 -t"
    # str_buttontext - button text in servers popup window
    servers_nested_list_data = vnjsonstructure['1 линия']
    # servers popup window size, different for each line
    serverswnd_width = vnjsonstructure['serverspopup rect 1 line'][0]
    serverswnd_height = vnjsonstructure['serverspopup rect 1 line'][1]
    # assign bitmap 16x16 for server status not connected, different png for each line
    icon_connected = icon_connected1
    icon_notconnected = icon_notconnected1
    # assign 70 bimtaps png for chosen server
    allservers_bitmap = autopy.bitmap.Bitmap.open(line1_allservers_png)
# the same for line2
elif autopy.bitmap.capture_screen().count_of_bitmap(autopy.bitmap.Bitmap.open(line2_detector_png),
                                                    rect=((1, 83), (634, 174))) > 0:
    servers_nested_list_data = vnjsonstructure['2 линия']
    serverswnd_width = vnjsonstructure['serverspopup rect 2 line'][0]
    serverswnd_height = vnjsonstructure['serverspopup rect 2 line'][1]
    icon_connected = icon_connected2
    icon_notconnected = icon_notconnected2
    allservers_bitmap = autopy.bitmap.Bitmap.open(line2_allservers_png)
# when videonet configuration window's area with academ1 or avtozav1 is not visible, criterror
else:
    raise Exception("Перекрыта область сканирования")
# making lists for purposes
# each list are [x_pos_list, [stations_list, ip_list, btntxt_list]] indexes from json's nested list
xpos_list = []
stations_list = []
ip_list = []
buttontext_list = []
valid_xpos_list = []
# getting every second index from nested list [0, [list], 2, [list],...]
for i in range(len(servers_nested_list_data)):
    if i % 2 == 0:
        xpos_list.append(servers_nested_list_data[i])
        stations_list.append(servers_nested_list_data[i + 1][0])
        ip_list.append(servers_nested_list_data[i + 1][1])
        buttontext_list.append(servers_nested_list_data[i + 1][2])
# getting int png_list_length = number of servers from json for use in main script for "for cycle"
# to detect bitmap icon_notconnected.png 
for j in range(len(servers_nested_list_data)):
    if j % 2 == 0 and servers_nested_list_data[j] != -1:
        valid_xpos_list.append(servers_nested_list_data[j])
valid_xpos_list_length = len(valid_xpos_list)
# all legal servers from json, -1 value excluded, combining to dictionary {x_pos : True}
# for script that put False value if it's index in bitmap of 70 bitmaps is a server that is not connected
valid_xpos_bool_dict = dict.fromkeys(valid_xpos_list, True)
# servers popup rows value is number of full filled rows of buttons, divided as int
rows = round(len(buttontext_list) / 5)
# divided as rational number when last row has less than 5 buttons, not int, add 1
if round((len(buttontext_list) / 5), 1) - rows != 0:
    rows += 1
# number of columns of buttons in servers popup is constant
cols = 5
# configuring button id for checkbox windows. id is formatted 3iijj 
# where 3 is for checkboxes ii: 01-99 rows, jj: 01-99 columns
checkbox_butid_list = []
index_stations_list = 0
for i in range(rows):
    i += 1
    for j in range(cols):
        j += 1
        if index_stations_list < len(buttontext_list):
            button_ij = 30000 + i * 100 + j
            checkbox_butid_list.append(button_ij)
            index_stations_list += 1
# make a dictionary {button id : x position in 70 bitmaps}
checkboxbutid_xpos_dict = dict(zip(checkbox_butid_list, xpos_list))
# time interval in seconds how often to refresh content in main window, if it is changed
app_window_refresher_thread_interval = vnjsonstructure["app_window_refresher_thread_interval"]
# time interval in seconds how often to scan screen for bitmap icon_notconnected
screen_bitmap_scanner_thread_interval = vnjsonstructure["screen_bitmap_scanner_thread_interval"]
# servers popup position x=left, y=top
serverswnd_x_pos = vnjsonstructure["pos x serverslist wnd"]
serverswnd_y_pos = vnjsonstructure["pos y serverslist wnd"]
# cmd shell window position x=left, y=top
cmdwnd_x_pos = vnjsonstructure["pos x cmd launch wnd"]
cmdwnd_y_pos = vnjsonstructure["pos y cmd launch wnd"]
# cmd shell window colors blue background, light text
cmd_wnd_color = vnjsonstructure["cmd window color"]
# buttions: checkbox width, button width
checkbox_btn_width = vnjsonstructure["checkbox button width"]
cmd_btn_width = vnjsonstructure["cmd button width"]
# buttons height
buttons_height = vnjsonstructure["buttons height"]
# margins to calculate spaces between buttons and groupboxes
buttons_margin_x = vnjsonstructure["buttons margin x"]
buttons_margin_y = vnjsonstructure["buttons margin y"]
# cmd popup radiobutton click area
radio_btn_width = vnjsonstructure["radiobuttons width"]
radio_btn_height = vnjsonstructure["radiobuttons height"]
# checked radiobutton: row 2 col 2
default_radiobutton_butid = vnjsonstructure["default radiobutton"]


class TBBUTTON(ctypes.Structure):
    """class to pack TBBUTTON structure, python has not this structure"""
    _pack_ = 1
    _fields_ = [
        ('iBitmap', ctypes.c_int),
        ('idCommand', ctypes.c_int),
        ('fsState', ctypes.c_ubyte),
        ('fsStyle', ctypes.c_ubyte),
        ('bReserved', ctypes.c_ubyte * 2),
        ('dwData', ctypes.c_ulong),
        ('iString', ctypes.c_int)
    ]


class CreateCmdProcess():
    """launch popup cmd shell with ping ip -t"""

    def __init__(self, station="", ip="", exec="", pwrcfgcmdtitle=""):
        super(CreateCmdProcess, self).__init__()
        self.pwrcfgcmdtitle = pwrcfgcmdtitle
        self.exec = exec
        self.station = station
        self.ip = ip

    def createwnd(self, station, ip):
        # https://ss64.com/nt/cmd.html
        # structure of cmd shell variables
        cmdsettings = win32process.STARTUPINFO()
        # set not default appearance of window
        cmdsettings.dwFlags = win32con.STARTF_USEPOSITION
        # get x, y from json
        cmdsettings.dwX = cmdwnd_x_pos
        cmdsettings.dwY = cmdwnd_y_pos
        return win32process.CreateProcess(
            "C:\\Windows\\System32\\cmd.exe",
            "cmd.exe /k COLOR " + cmd_wnd_color + " && TITLE " + station + "&& PING " + ip,
            # 0A 1B https://ss64.com/nt/color.html
            None, None, False, win32con.CREATE_NEW_CONSOLE, None, None, cmdsettings)

    def set_power_configuration(self, commandstringexec, pwrcfgcmdtitle):
        cmdsettings = win32process.STARTUPINFO()
        cmdsettings.dwX = cmdwnd_x_pos
        cmdsettings.dwY = cmdwnd_y_pos
        cmdsettings.lpTitle = pwrcfgcmdtitle
        return win32process.CreateProcess(
            "C:\\Windows\\System32\\cmd.exe",
            "cmd.exe /c COLOR " + cmd_wnd_color + " && " + commandstringexec,
            # 0A 1B https://ss64.com/nt/color.html
            None, None, False, (win32con.CREATE_NEW_CONSOLE | win32con.CREATE_NO_WINDOW), None, None, cmdsettings)



class ToolbarButton():
    """main window "toolbar" buttons Clock, Cmd, Servers functionality if buttons pressed"""

    def __init__(self, title="", station="", ip="", hWndToolbar="", tbbut_id=""):
        super(ToolbarButton, self).__init__()
        self.tbbut_id = tbbut_id
        self.title = title
        self.station = station
        self.ip = ip
        self.hWndToolbar = hWndToolbar


    def store_hWndToolbar(self, hWndToolbar):
        global hWndToolbar_value
        hWndToolbar_value = hWndToolbar
        # print("i am in ToolbarButton().store_hWndToolbar(hWndToolbar) hWndToolbar =", hWndToolbar_value)

    def toggledaynight(self, tbbut_id):
        toolbar_button_state = win32gui.SendMessage(hWndToolbar_value, commctrl.TB_GETSTATE, tbbut_id, 0)

        # if toolbar_button_state == 5:
        #    print ("pressed")
        # elif toolbar_button_state == 4:
        #    print("freed")

        if toolbar_button_state == commctrl.TBSTATE_ENABLED:
            # balanced
            win32gui.SendMessage(hWndToolbar_value, commctrl.TB_CHANGEBITMAP, 14, 4)
            CreateCmdProcess().set_power_configuration(commandstringexec='powercfg /s 381b4222-f694-41f0-9685-ff5bb260df2e', pwrcfgcmdtitle ='balanced')
        elif toolbar_button_state == (commctrl.TBSTATE_CHECKED | commctrl.TBSTATE_ENABLED):
            # economy
            win32gui.SendMessage(hWndToolbar_value, commctrl.TB_CHANGEBITMAP, 14, 5)
            CreateCmdProcess().set_power_configuration(commandstringexec='powercfg /s a1841308-3541-4fab-bc81-f71556f20b4a', pwrcfgcmdtitle='economy')


        """
            def get_hWndToolbar(self):
                return hWndToolbar_value
            # hwnd_toolbar = ToolbarButton().get_hWndToolbar()
                        
            bk = win32gui.SendMessage(hwnd_toolbar, commctrl.TB_BUTTONCOUNT , 0, 0)
            # print(bk)

            # https://docs.microsoft.com/en-us/windows/win32/controls/tb-getbutton

            # win32gui.SendMessage(hwnd_toolbar, commctrl.TB_PRESSBUTTON, 13, 1)

            gs = win32gui.SendMessage(hwnd_toolbar, commctrl.TB_GETSTATE, 14, 0)

            #print(gs) # 5 or 4

            #if gs == 5:
            #    print ("pressed")
            #elif gs == 4:
            #    print("freed")

            if gs == commctrl.TBSTATE_ENABLED:
                print("freed")
            elif gs == (commctrl.TBSTATE_CHECKED | commctrl.TBSTATE_ENABLED):
                print ("pressed")

            #print(commctrl.TBSTATE_CHECKED)
            #print(commctrl.TBSTATE_ELLIPSES)
            #print(commctrl.TBSTATE_ENABLED)
            #print(commctrl.TBSTATE_HIDDEN)
            #print(commctrl.TBSTATE_INDETERMINATE)
            #print(commctrl.TBSTATE_MARKED)
            #print(commctrl.TBSTATE_PRESSED)
            #print(commctrl.TBSTATE_WRAP)

        """

    # toggle hide/show windows Clock, Servers
    def toggleshow(self, title):
        hwnd = win32ui.FindWindow(None, title)
        if hwnd.IsWindowVisible():
            hwnd.ShowWindow(win32con.SW_HIDE)
        else:
            hwnd.ShowWindow(win32con.SW_SHOW)

    # button Cmd launches first not connected server in main window or do nothing if OK
    def fastping(self):
        if displayinfo != "" and displayinfo != "OK":
            station = displayinfo.split("\n", 1)[0]
            ind = stations_list.index(station)
            ip = ip_list[ind]
            CreateCmdProcess().createwnd(station, ip)




class ClockPopup():
    """Clock popup window: create window. create buttons, close app function"""

    def __init__(self, title="", hWnd_q="", wndClassAtom="", qwidth="", qheight="", hWndApp="", hInstance=""):
        super(ClockPopup, self).__init__()
        self.title = title
        self.hWnd_q = hWnd_q
        self.wndClassAtom = wndClassAtom
        self.hWndApp = hWndApp
        self.hInstance = hInstance
        self.qwidth = qwidth
        self.qheight = qheight

    # gracefully quit application
    def app_terminate(self):
        ctypes.windll.user32.PostMessageA(win32gui.FindWindow(0, main_window_title), win32con.WM_CLOSE, 0, 0)

    # create Clock popup window
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

    # when Clock popup window is being moved and then is being hidden,
    # the next time it will be shown anchored to the top left corner of main window
    # removed because of different relative coordinates for windows schemes
    def changewndshowposition(self, title):
        pass
        # hwnd = win32ui.FindWindow(None, title)
        # (wndleft, wndtop, wndright, wndbottom) = hwnd.GetClientRect()
        # hWndApp = win32ui.FindWindow(None, main_window_title)
        # (hleft, htop, hright, hbottom) = hWndApp.GetWindowRect()
        # hwnd.MoveWindow((hleft, htop-wndbottom-39, hleft+wndright+16, htop), 1)

        # radiobuttons creation

    def addradiobutton(self, title, hWnd_q, hInstance, row, col, butID):
        hWndradiobutton = win32gui.CreateWindow(
            'BUTTON', title,
            win32con.WS_VISIBLE | win32con.WS_CHILD | win32con.BS_LEFT | win32con.BS_AUTORADIOBUTTON,
            buttons_margin_x * 2 + buttons_margin_x * 2 * col + radio_btn_width * (col - 1),
            buttons_margin_y * 2 + buttons_margin_y * 2 * row + radio_btn_height * (row - 1) + buttons_margin_y * 2,
            radio_btn_width, radio_btn_height,
            hWnd_q, butID, hInstance, None)
        return hWndradiobutton

    # groupbox creation
    def addgroupbox(self, qwidth, hWnd_q, hInstance, butID):
        hWndgroupbox = win32gui.CreateWindow(
            'BUTTON', None,
            win32con.WS_VISIBLE | win32con.WS_CHILD | win32con.BS_CENTER | win32con.BS_GROUPBOX | win32con.BS_BOTTOM,
            buttons_margin_x * 2, buttons_margin_y * 2, qwidth - buttons_margin_x * 8,
            buttons_margin_y * 8 + radio_btn_height * 2 + buttons_margin_y * 2, hWnd_q,
            20001, hInstance, None)
        return hWndgroupbox

    # close button creation
    def addclosebutton(self, title, qwidth, qheight, hWnd_q, hInstance, butID):
        hWndclosebutton = win32gui.CreateWindow("BUTTON", "Закрыть",
                                                win32con.WS_VISIBLE | win32con.WS_CHILD | win32con.BS_CENTER | win32con.BS_DEFPUSHBUTTON,
                                                round(qwidth / 2) - round(radio_btn_width / 2),
                                                qheight - round(radio_btn_height / 2) - round((qheight - (
                                                            buttons_margin_y * 2 + radio_btn_height * 2 + buttons_margin_y * 2)) / 2),
                                                radio_btn_width, radio_btn_height, hWnd_q, 20000, hInstance, None)
        return hWndclosebutton


class CmdPopup():
    """Cmd popup button #2 behaviour: launch CMD.exe and PING failed server"""

    def __init__(self, wParam=""):
        super(CmdPopup, self).__init__()
        self.wParam = wParam

    # create cmd shell window with ping
    def cmdlaunch(self, wParam):
        # get an index in list in json from wParam sctucture wParam=40302 row=3, col=2 index = (03-1)*5 + (02-1) = 11
        station_index = ((int(repr(wParam)[1:3])) - 1) * 5 + (int(repr(wParam)[3:5])) - 1
        # assign string value for titte in cmd shell window
        station = stations_list[station_index]
        # and value for ping "10.99.12.1 -t"
        ip = ip_list[station_index]
        # create cmd shell window       
        CreateCmdProcess().createwnd(station, ip)


class ServersPopup():
    """Servers popup "toolbar" button #3 functionality: create window, create buttons"""

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

        # create Servers popup window

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
            serverswnd_x_pos,  # x
            serverswnd_y_pos,  # y
            swidth,  # nWidth
            sheight,  # nHeight
            hWndApp,  # hWndParent
            0,  # hMenu
            hInstance,  # hInstance
            None  # lpParam
        )
        return hWndserverspopup

    # create two buttons: checkbox (ignore fail detection) and pushbutton (launch cmd shell)
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
    """main script to detect info to display in main window: colored text"""

    def __init__(self):
        super(TextOnScreen, self).__init__()

    # scan screen for not_connected bitmap and set values for color and text to display 
    def scan(self):
        # get screenshot
        screen_picture = autopy.bitmap.capture_screen()
        # each found bitmap as (x, y) are added to tuple, rect= is for a region scan, not entire screenshot
        notconnected_positions_tuple = screen_picture.find_every_bitmap(needle=icon_notconnected,
                                                                        tolerance=.1)  # , rect=((1, 83), (770, 174))
        # for testing purpose: connected_positions_tuple = screen_picture.find_every_bitmap(icon_connected, rect=((1, 83), (634, 174))) 
        # if have not found a bitmap with icon_notconnected (tuple len=0) -> set "OK" black
        if len(notconnected_positions_tuple) == 0:
            displayinfo = "OK"
            txtcolor = black_color
        # if bitmaps have been found -> set multiline text, color red
        else:
            displayinfo = ""
            txtcolor = red_color
            # get every (x,y) from tuple
            for crop_nc in range(len(notconnected_positions_tuple)):
                # get a rect region of 58x16 bitmap to the right hand from servericon in the screenshot
                x1cn, y1cn, x2cn, y2cn = notconnected_positions_tuple[crop_nc][0] + 16, \
                                         notconnected_positions_tuple[crop_nc][1], 58, 16
                cropped_58x16_serv_notcon = screen_picture.cropped(rect=((x1cn, y1cn), (x2cn, y2cn)))
                # png_list is a list of x_pos of bitmaps with servername like academ1.png
                for xpos in valid_xpos_list:
                    # go through ~70 bitmaps png file to find x_pos bitmap, blank bitmaps are ommited because of png_list structure
                    server_iterate_bitmap = allservers_bitmap.cropped(rect=((xpos, 0), (58, 16)))
                    # when bitmap from screen_picture(after cut from screenshot with tuple x,y and 58x16)
                    # is equal to servers_picture(70 bitmaps x_pos indexed) and in dictionary
                    # servers_list_dict_from_name_bool value on key=x_pos is set True(do not ignore failed server)
                    if cropped_58x16_serv_notcon.is_bitmap_equal(server_iterate_bitmap, 0.1) and \
                            valid_xpos_bool_dict[xpos]:
                        # add string to text to display in main window
                        displayinfo += servers_nested_list_data[servers_nested_list_data.index(xpos) + 1][0] + "\n"
            # text color is set to red but text is "OK" because if not_connected bitmaps were found at least 1 time but
            # servers_list_dict_from_name_bool on x_pos is set False because of checkbox is checked
            if displayinfo == "":
                displayinfo = "OK"
                txtcolor = red_color
        # two values returned on TextOnScreen().scan() request
        return displayinfo, txtcolor


# main script for winapi creating windows and threads
def main():
    """main class"""
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
    win32gui.SetWindowPos(hWndApp, win32con.HWND_TOPMOST, appwnd_left, appwnd_top, appwnd_right, appwnd_bottom, win32con.SWP_SHOWWINDOW)
    win32gui.ShowWindow(hWndApp, win32con.SW_SHOWNORMAL)

    """**********************************************************************
    *              CLOCK WINDOW INIT - Toolbar button 2
    **********************************************************************"""
    # qwidth, qheight = 300, 200
    hWnd_q = ClockPopup().createwnd(hWndApp, clockwnd_width, clockwnd_height, wndClassAtom, hInstance)
    hWndButton20101 = ClockPopup().addradiobutton("7:50 / 19:50", hWnd_q, hInstance, row=1, col=1, butID=20101)
    hWndButton20102 = ClockPopup().addradiobutton("7:50", hWnd_q, hInstance, row=1, col=2, butID=20102)
    hWndButton20201 = ClockPopup().addradiobutton("19:50", hWnd_q, hInstance, row=2, col=1, butID=20201)
    hWndButton20202 = ClockPopup().addradiobutton("Не выключать", hWnd_q, hInstance, row=2, col=2, butID=20202)
    hWndGrpBox20001 = ClockPopup().addgroupbox(clockwnd_width, hWnd_q, hInstance, butID=20001)
    hWndButton20000 = ClockPopup().addclosebutton("Закрыть", clockwnd_width, clockwnd_height, hWnd_q, hInstance, butID=20000)
    check_default_radiobutton_clockpopup = win32ui.FindWindow(None, clock_popup_title)
    check_default_radiobutton_clockpopup.CheckRadioButton(20101, 20202, default_radiobutton_butid)

    """**********************************************************************
    *              SERVERS WINDOW INIT - Toolbar button 3
    **********************************************************************"""
    hWnd_s = ServersPopup().createwnd(serverswnd_width, serverswnd_height, hWndApp, wndClassAtom, hInstance)
    # calculating row count for columns=5
    rows = round(len(buttontext_list) / 5)
    if round((len(buttontext_list) / 5), 1) - rows != 0:
        rows += 1
    cols = 5
    # buttons drawing
    # if a button has not value, it is not shown
    index_stations_list = 0
    for i in range(rows):
        i += 1
        for j in range(cols):
            j += 1
            if index_stations_list < len(buttontext_list):
                # print(index_stations_list)
                button_ij = 30000 + i * 100 + j
                # print(btntxt_list[index_stations_list])
                # print(button_ij, btntxt_list[index_stations_list], index_stations_list , "i=", i, "j=",j)
                hWndcheckboxbutton, hwndcmdbutton = ServersPopup().addcheckboxcmdbuttons(
                    buttontext_list[index_stations_list], hWnd_s, hInstance, row=i, col=j, butID=button_ij)
                if buttontext_list[index_stations_list] == "":
                    win32gui.ShowWindow(hWndcheckboxbutton, win32con.SW_HIDE)
                    win32gui.ShowWindow(hwndcmdbutton, win32con.SW_HIDE)
                index_stations_list += 1
    hWndGrpBox30001 = ServersPopup().addgroupbox(serverswnd_width, serverswnd_height, hWndApp, hWnd_s, hInstance, butID=30001)
    hWndButton30000 = ServersPopup().addclosebutton(serverswnd_width, serverswnd_height, hWndApp, "Закрыть", hWnd_s, hInstance, butID=30000)

    """**********************************************************************
    *              TOOLBAR INIT
    **********************************************************************"""
    hImageList = None
    numButtons = 1
    # bitmapSize_x = 45
    # bitmapSize_y = 45
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
    hImageList = win32gui.ImageList_Create(toolbar_button_size_x, toolbar_button_size_y, win32gui.ILC_COLOR24 | win32gui.ILC_MASK,
                                           numButtons, 0)
    win32gui.SendMessage(hWndToolbar, commctrl.TB_SETIMAGELIST, 0, hImageList)
    win32gui.SendMessage(hWndToolbar, commctrl.TB_SETBUTTONSIZE, 0, win32api.MAKELONG(toolbar_button_size_x, toolbar_button_size_y))
    tbstring = win32gui.SendMessage(hWndToolbar, commctrl.TB_ADDSTRING, hInstance, 33)
    tbButtons = TBBUTTON()
    tbButtons.bReserved = 0, 0
    tbButtons.dwData = 0
    tbButtons.iString = tbstring
    # Toolbar button 1
    hBitmap = win32gui.LoadImage(0, toolbar_image_path["clock"], win32gui.IMAGE_BITMAP, toolbar_button_size_x, toolbar_button_size_y,
                                 win32gui.LR_LOADFROMFILE)
    toolbarBitmap = win32gui.ImageList_Add(hImageList, hBitmap, 0)
    tbButtons.iBitmap = toolbarBitmap
    tbButtons.idCommand = 11
    tbButtons.fsState = commctrl.TBSTATE_ENABLED | commctrl.TBSTATE_WRAP
    tbButtons.fsStyle = commctrl.BTNS_BUTTON  # | commctrl.BTNS_CHECK
    win32gui.SendMessage(hWndToolbar, commctrl.TB_ADDBUTTONS, 1, tbButtons)
    # Toolbar button 2
    hBitmap = win32gui.LoadImage(0, toolbar_image_path["cmd"], win32gui.IMAGE_BITMAP, toolbar_button_size_x, toolbar_button_size_y,
                                 win32gui.LR_LOADFROMFILE)
    toolbarBitmap = win32gui.ImageList_Add(hImageList, hBitmap, 0)
    tbButtons.iBitmap = toolbarBitmap
    tbButtons.idCommand = 12
    tbButtons.fsState = commctrl.TBSTATE_ENABLED | commctrl.TBSTATE_WRAP
    tbButtons.fsStyle = commctrl.BTNS_BUTTON
    win32gui.SendMessage(hWndToolbar, commctrl.TB_ADDBUTTONS, 1, tbButtons)
    # Toolbar button 3
    hBitmap = win32gui.LoadImage(0, toolbar_image_path["servers"], win32gui.IMAGE_BITMAP, toolbar_button_size_x, toolbar_button_size_y,
                                 win32gui.LR_LOADFROMFILE)
    toolbarBitmap = win32gui.ImageList_Add(hImageList, hBitmap, 0)
    tbButtons.iBitmap = toolbarBitmap
    tbButtons.idCommand = 13
    tbButtons.fsState = commctrl.TBSTATE_ENABLED | commctrl.TBSTATE_WRAP
    tbButtons.fsStyle = commctrl.BTNS_BUTTON
    win32gui.SendMessage(hWndToolbar, commctrl.TB_ADDBUTTONS, 1, tbButtons)
    # Toolbar button 4
    hBitmap = win32gui.LoadImage(0, toolbar_image_path["powercfg_default"], win32gui.IMAGE_BITMAP, toolbar_button_size_x,
                                 toolbar_button_size_y,
                                 win32gui.LR_LOADFROMFILE)
    toolbarBitmap = win32gui.ImageList_Add(hImageList, hBitmap, 0)
    tbButtons.iBitmap = toolbarBitmap
    tbButtons.idCommand = 14
    tbButtons.fsState = commctrl.TBSTATE_ENABLED | commctrl.TBSTATE_WRAP
    tbButtons.fsStyle = commctrl.BTNS_CHECK
    win32gui.SendMessage(hWndToolbar, commctrl.TB_ADDBUTTONS, 1, tbButtons)

    hBitmap = win32gui.LoadImage(0, toolbar_image_path["powercfg_day"], win32gui.IMAGE_BITMAP,
                                 toolbar_button_size_x,
                                 toolbar_button_size_y,
                                 win32gui.LR_LOADFROMFILE)
    win32gui.ImageList_Add(hImageList, hBitmap, 0)
    hBitmap = win32gui.LoadImage(0, toolbar_image_path["powercfg_night"], win32gui.IMAGE_BITMAP,
                                 toolbar_button_size_x,
                                 toolbar_button_size_y,
                                 win32gui.LR_LOADFROMFILE)
    win32gui.ImageList_Add(hImageList, hBitmap, 0)

    win32gui.ShowWindow(hWndToolbar, win32con.SW_SHOWNORMAL)

    ToolbarButton().store_hWndToolbar(hWndToolbar=hWndToolbar)

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


# thread checks for button id checked and terminates app at the time
# time offset is -3 because of Windows specific timezone settings
def app_terminator_scheduler_thread(hWndApp):
    global schedule_time_clock_popup_radiobutton_id
    schedule_time_clock_popup_radiobutton_id = default_radiobutton_butid
    while True:
        # tm_year=2020, tm_mon=6, tm_mday=20, tm_hour=19, tm_min=11, tm_sec=3, tm_wday=5, tm_yday=172, tm_isdst=0
        if schedule_time_clock_popup_radiobutton_id == 20101:  # 7:50 / 19:50
            if time.localtime().tm_hour == 16 or time.localtime().tm_hour == 4:
                if time.localtime().tm_min == 50:
                    ClockPopup().app_terminate()
        if schedule_time_clock_popup_radiobutton_id == 20102:  # 7:50
            if time.localtime().tm_hour == 4 and time.localtime().tm_min == 50:
                ClockPopup().app_terminate()
        if schedule_time_clock_popup_radiobutton_id == 20201:  # 19:50
            if time.localtime().tm_hour == 16 and time.localtime().tm_min == 50:
                ClockPopup().app_terminate()
        if schedule_time_clock_popup_radiobutton_id == 20202:  # Не выключать
            time.sleep(10)
        time.sleep(10)


# constantly scanning screen, if variables displayinfo or textcolor has changed
# then set flag refreshercommander false and cause main app refresher
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


# checking flag refreshercommander if variables displayinfo or textcolor has changed
# true: do refresh main app window, false: do nothing 
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
            ClockPopup().changewndshowposition(title=clock_popup_title)
            ToolbarButton().toggleshow(title=clock_popup_title)
        # Toolbar button [2] CMD when pressed
        if wParam == 12:
            ToolbarButton().fastping()
        # Toolbar button [3] SERVERS when pressed
        if wParam == 13:
            ToolbarButton().toggleshow(title=servers_popup_title)
        # Toolbar button [3] SERVERS when pressed
        if wParam == 14:
            ToolbarButton().toggledaynight(tbbut_id=14)

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
                # "\n"+repr(butID+10000)
                # print("xpos=", checkbox_dict_from_wparam_xpos.get(wParam))
                # servers_list_dict_from_name_bool.pop(checkbox_dict_from_wparam_xpos.get(wParam))
                # servers_list_dict_from_name_bool.update({checkbox_dict_from_wparam_xpos.get(wParam):False})
                # print(servers_list_dict_from_name_bool)
                # print(wParam, "checked") # 3rrcc
            if radiobuttonStatus == win32con.BST_UNCHECKED:
                schedule_time_clock_popup_radiobutton_id = wParam
                # servers_list_dict_from_name_bool.pop(checkbox_dict_from_wparam_xpos.get(wParam))
                # servers_list_dict_from_name_bool.update({checkbox_dict_from_wparam_xpos.get(wParam):True})
                # print(wParam, "unchecked")

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
            global valid_xpos_bool_dict
            # if checkbox checked: pop and add key, value changed to false
            if checkboxStatus == win32con.BST_CHECKED:
                # "\n"+repr(butID+10000)
                # print("xpos=", checkbox_dict_from_wparam_xpos.get(wParam))
                valid_xpos_bool_dict.pop(checkboxbutid_xpos_dict.get(wParam))
                valid_xpos_bool_dict.update({checkboxbutid_xpos_dict.get(wParam): False})
                # print(servers_list_dict_from_name_bool)
                # print(wParam, "checked") # 3rrcc
            # if checkbox unchecked: pop and add key, value changed to true
            if checkboxStatus == win32con.BST_UNCHECKED:
                valid_xpos_bool_dict.pop(checkboxbutid_xpos_dict.get(wParam))
                valid_xpos_bool_dict.update({checkboxbutid_xpos_dict.get(wParam): True})
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
        fontSize = font_size
        lf = win32gui.LOGFONT()
        lf.lfFaceName = font_name
        lf.lfHeight = int(round(dpiScale * fontSize))
        lf.lfWeight = font_weight
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
    # here is GW_OWNER flag but actually it means OWNED
    # complex comparison made to hide exact child windows and to block WM_DESTROY command message
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
