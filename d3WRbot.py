# simple diablo 3 warriors rest pixel bot with a little GUI in wxPython
# only works in windowed mode with a resolution of 1064x758

import win32api
import win32con
import win32gui
import win32com.client
import ImageGrab
import os
import time
import wx
import threading
import sys

# observer pattern
class Observable(list):

    def addObserver(self, observer):
        self.append(observer)

    def notifyObservers(self, event):
        for observer in self:
            observer.update(event)


class Observer(object):

    def __init__(self, name):
        self.name = name

    def update(self, event):
        pass


class DiabloBot(Observable):
    def __init__(self):
        self.refreshDiabloWindow()
        self.shell = win32com.client.Dispatch("WScript.Shell")
        self.legendarycount = 0
        self.rarecount = 0
        self.setcount = 0
        self.runcount = 0

    def refreshDiabloWindow(self):
        diablowindow = win32gui.FindWindow(None, "Diablo III")
        if (diablowindow != 0):
            diablorectangle = win32gui.GetWindowRect(diablowindow)
            # setting padding to corners of diablo window
            # grab screen with x1 and y1 minus one because the rectangle are exclusive:
            #(right,bottom) lies immediately outside the rectangle.
            # we need it to be inside.
            self.x_2 = diablorectangle[0]
            self.y_2 = diablorectangle[1]
            self.x_1 = diablorectangle[2] - 1
            self.y_1 = diablorectangle[3] - 1
        else:
            print "could not locate diablo window - did you start diablo yet?"

    gamecoords = {'moveright_click': (375, 234),
                  'leavegame_click': (533, 406),
                  'resumegame_click': (160, 300),
                  'targetself_click': (530, 315),
                  'rightofself_click': (800, 350),
                  'last_inventoryspace_coord': (975, 550),
                  'targetleftofself_click': (560, 315),
                  'chest_click': (595, 200),
                  'last_tab_cord': (300, 500),
                  'tp_click': (500, 325),
                  'repairsymbol': (785, 45),
                  'repair_tab': (347, 352),
                  'repair_all': (214, 421),
                  'first_tab': (350, 180),
                  'second_tab': (350, 260),
                  'third_tab': (350, 350),
                  'death_check': (447, 595)}

    gamecolors = {'inventory_empty_block': (19, 11, 5),
                  'rare_itemcolor': (255, 255, 0),
                  'legendary_itemcolor': (191, 100, 47),
                  'set_itemcolor': (0, 255, 0),
                  'red_repaircolor': (222, 0, 0),
                  'revive_textcolor': (218, 148, 74)}

    def screenGrab(self):
        box = (self.x_2, self.y_2, self.x_1, self.y_1)
        im = ImageGrab.grab(box)
        return im

    def screenGrabToFile(self, name):
        box = (self.x_2, self.y_2, self.x_1, self.y_1)
        im = ImageGrab.grab(box)
        im.save(os.getcwd() + "\\" + name + str(int(time.time())) + '.png', 'PNG')

    def leftClick(self):
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        time.sleep(.1)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

    def rightClick(self):
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0)
        time.sleep(.1)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0)

    def rightDown(self):
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0)
        time.sleep(.1)

    def rightUp(self):
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0)
        time.sleep(.1)

    def leftDown(self):
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        time.sleep(.1)

    def leftUp(self):
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
        time.sleep(.1)

    def mousePos(self, cord):
        win32api.SetCursorPos((self.x_2 + cord[0], self.y_2 + cord[1]))

    def getCords(self):
        x, y = win32api.GetCursorPos()
        x = x - self.x_2
        y = y - self.y_2
        print x, y

    def buff(self):
        self.shell.SendKeys("3")
        self.shell.SendKeys("4")

    def repair(self):
        print "repairing"
        self.mousePos((1004, 43))
        time.sleep(0.3)
        self.leftClick()
        time.sleep(3.5)
        self.mousePos((492, 170))
        time.sleep(0.3)
        self.leftClick()
        time.sleep(3)
        self.mousePos(self.gamecoords['repair_tab'])
        time.sleep(0.3)
        self.leftClick()
        self.mousePos(self.gamecoords['repair_all'])
        time.sleep(0.3)
        self.leftClick()
        self.shell.SendKeys("I")
        self.mousePos((389, 633))
        time.sleep(0.3)
        self.leftClick()
        time.sleep(2)
        self.mousePos((389, 633))
        time.sleep(0.3)
        self.leftClick()
        time.sleep(2)
        self.mousePos((49, 406))
        time.sleep(0.3)
        self.leftClick()
        time.sleep(1.5)

    def needRepair(self):
        image = self.screenGrab()
        if (image.getpixel(self.gamecoords['repairsymbol']) == self.gamecolors['red_repaircolor']):
            return True
        else:
            return False

    def pickupRareItems(self):
        if self.isDead():
            return False
        image = self.screenGrab()
        for x in range(1, self.x_1 - self.x_2):
            # y from 100 to 75% of max is the area where items are dropping
            for y in range(100, self.y_1 / 100 * 75):
                if (image.getpixel((x, y)) == self.gamecolors['rare_itemcolor']):
                    self.mousePos((x, y))
                    self.leftClick()
                    time.sleep(2.5)
                    self.rarecount = self.rarecount + 1
                    self.notifyObservers("rarecount")
                    self.pickupRareItems()
                    return True
        return False

    def pickupLegendaryItems(self):
        image = self.screenGrab()
        for x in range(1, self.x_1 - self.x_2):
            # y from 100 to 75% of max is the area where items are dropping
            for y in range(100, self.y_1 / 100 * 75):
                if (image.getpixel((x, y)) == self.gamecolors['legendary_itemcolor']):
                    print "legendary item found!"
                    self.screenGrabToFile("legendary")
                    time.sleep(2.5)
                    if self.isDead():
                        sys.exit(1)
                    self.mousePos((x, y))
                    self.leftClick()
                    time.sleep(2.5)
                    self.legendarycount = self.legendarycount + 1
                    self.notifyObservers("legendarycount")
                    self.pickupLegendaryItems()
                    return True
        return False

    def pickupSetItems(self):
        image = self.screenGrab()
        for x in range(1, self.x_1 - self.x_2):
            # y from 100 to 75% of max is the area where items are dropping
            for y in range(100, self.y_1 / 100 * 75):
                if (image.getpixel((x, y)) == self.gamecolors['set_itemcolor']):
                    print "set item found!"
                    self.screenGrabToFile("set")
                    time.sleep(2.5)
                    if self.isDead():
                        sys.exit(1)
                    self.mousePos((x, y))
                    self.leftClick()
                    time.sleep(2.5)
                    self.setcount = self.setcount + 1
                    self.notifyObservers("setcount")
                    self.pickupSetItems()
                    return True
        return False

    def exitDiablo(self):
        print "exiting diablo"
        self.shell.sendKeys("{Esc}")
        self.mousePos(self.gamecoords['leavegame_click'])
        self.leftClick()
        time.sleep(10)

    def resumeDiablo(self):
        self.mousePos(self.gamecoords['resumegame_click'])
        self.leftClick()
        time.sleep(3)

    def castBlizzard(self):
        self.rightClick()

    def castHydra(self):
        self.shell.SendKeys("2")

    def castDiamondSkin(self):
        self.shell.SendKeys("1")

    def castSeekerMissile(self):
        win32api.keybd_event(win32con.VK_LSHIFT, 0, win32con.KEYEVENTF_EXTENDEDKEY | 0, 0)
        self.leftDown()
        time.sleep(2.5)
        self.leftUp()
        win32api.keybd_event(win32con.VK_LSHIFT, 0, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)

    def isInventoryFull(self):
        self.shell.SendKeys("I")
        time.sleep(1.5)
        image = self.screenGrab()
        if (image.getpixel(self.gamecoords['last_inventoryspace_coord']) != self.gamecolors['inventory_empty_block']):
            print "inventory is full"
            self.shell.SendKeys("i")
            return True
        else:
            self.shell.SendKeys("i")
            return False

    def isTabFull(self):
        image = self.screenGrab()
        if (image.getpixel(self.gamecoords['last_tab_cord']) != (21, 12, 8)):
            print "current tab full"
            return True
        else:
            return False

    def teleport(self):
        self.shell.SendKeys("T")
        time.sleep(8)

    def stashTabCheck(self):
        if (self.isTabFull()):
            self.moveToNextTab()
            return True
        return False

    def moveToStash(self):
        i = 1030
        j = 620
        self.mousePos((i, j))
        while(i > 760):
            while(j > 410):
                self.stashTabCheck()
                j = j - 32
                self.mousePos((i, j))
                self.rightClick()
            i = i - 32
            j = 620

    def isDead(self):
        image = self.screenGrab()
        if (image.getpixel(self.gamecoords['death_check']) == self.gamecolors['revive_textcolor']):
            return True
        else:
            return False

    def moveToNextTab(self):
        image = self.screenGrab()
        if (image.getpixel(self.gamecoords['first_tab']) == (85, 48, 24)):
            self.mousePos(self.gamecoords['second_tab'])
            time.sleep(0.5)
            self.leftClick()
        elif (image.getpixel(self.gamecoords['second_tab']) == (113, 75, 48)):
            self.mousePos(self.gamecoords['third_tab'])
            time.sleep(0.5)
            self.leftClick()
        elif (image.getpixel(self.gamecoords['third_tab']) == (74, 65, 55) and
            image.getpixel(self.gamecoords['last_inventoryspace_coord']) != self.gamecolors['inventory_empty_block']):
            print "no room left"
            sys.exit(1)

    def resetCounters(self):
        self.legendarycount = 0
        self.setcount = 0
        self.rarecount = 0
        self.runcount = 0
        self.notifyObservers("all")

    def enableItemText(self):
        self.shell.SendKeys("-")

    def runArchonBuild(self):
        # D-SKIN, HYDRA, MWEAPON, ENERGY ARMOR, MAGIC MISSILE, ARCHON - WITH BLUR, GALV. WARD, GLASS CANNON
        self.mousePos((630, 340))
        self.leftDown()
        time.sleep(7.4)
        self.leftUp()
        # cast hydra to the right of self + archon
        self.mousePos(self.gamecoords['rightofself_click'])
        self.castHydra()
        time.sleep(0.5)
        self.rightClick()
        time.sleep(0.5)
        self.castDiamondSkin()
        self.rightDown()
        for i in range(1, 15):
            self.mousePos((800, 280))
            time.sleep(0.5)
            self.mousePos((800, 350))
            time.sleep(0.5)
            self.mousePos((800, 480))
        self.rightUp()
        self.castDiamondSkin()

    def runBlizzBuild(self):
        # D-SKIN, HYDRA, MWEAPON, ENERGY ARMOR, MAGIC MISSILE, BLIZZARD - WITH BLUR, GALV. WARD, GLASS CANNON
         # run up to skeleton elite pack
        self.mousePos((630, 340))
        self.leftDown()
        time.sleep(7)
        self.leftUp()
        # cast hydra + blizzard to the right of self
        self.mousePos(self.gamecoords['rightofself_click'])
        self.castHydra()
        time.sleep(0.5)
        self.castBlizzard()
        self.castSeekerMissile()
        # cast diamondskin
        self.castDiamondSkin()
        self.castSeekerMissile()
        # move a bit left and recast blizzard + hydra on self
        self.mousePos((360, 355))
        self.leftClick()
        time.sleep(1)
        self.mousePos(self.gamecoords['targetleftofself_click'])
        self.castHydra()
        time.sleep(0.2)
        self.castBlizzard()
        self.castSeekerMissile()
        self.castSeekerMissile()
        # move a bit left and recast blizzard + hydra
        self.mousePos((360, 355))
        self.leftClick()
        time.sleep(1)
        self.mousePos(self.gamecoords['rightofself_click'])
        self.castHydra()
        self.mousePos(self.gamecoords['targetleftofself_click'])
        time.sleep(0.2)
        self.castBlizzard()
        self.mousePos(self.gamecoords['rightofself_click'])
        self.castSeekerMissile()
        self.castSeekerMissile()
        self.castSeekerMissile()
        self.castDiamondSkin()
        self.castBlizzard()
        self.castSeekerMissile()

    def runGame(self):
        time.sleep(2)
        # refresh diablo window coordinates
        self.refreshDiabloWindow()
        self.resumeDiablo()
        time.sleep(5)
        if (self.isInventoryFull()):
            self.teleport()
            # move to chest
            self.mousePos(self.gamecoords['chest_click'])
            time.sleep(0.3)
            self.leftClick()
            time.sleep(1.5)
            # stash items
            self.moveToStash()
            # move to tp
            self.mousePos(self.gamecoords['tp_click'])
            time.sleep(0.5)
            self.leftClick()
            time.sleep(1.5)
        if (self.needRepair()):
            print "repairing"
            self.teleport()
            self.repair()
        self.buff()
        # insert build
        self.runArchonBuild()
        self.mousePos(self.gamecoords['rightofself_click'])
        self.leftClick()
        time.sleep(2.6)
        self.pickupRareItems()
        self.pickupLegendaryItems()
        self.pickupSetItems()
        self.exitDiablo()
        time.sleep(10)


class WorkerThread(threading.Thread):
    def __init__(self, bot, numberofruns):
        threading.Thread.__init__(self)
        self.numberofruns = numberofruns
        self.bot = bot
        self._stop = threading.Event()
        self.start()

    def run(self):
        for i in range(self.numberofruns):
            self.bot.runGame()
            self.bot.runcount = i + 1
            self.bot.notifyObservers("runcount")
            if self.stopped():
                print "thread stopping"
                return

    def stop(self):
        self._stop.set()

    def stopped(self):
        return self._stop.isSet()


class BotGUI(wx.Frame, Observer):
    def __init__(self, parent, id):
        wx.Frame.__init__(self, parent, id, 'PandemoniumBot', size=(300, 200))
        Observer.__init__(self, 'BotGUI')
        panel = wx.Panel(self)
        startbutton = wx.Button(panel, label="Run game", pos=(10, 50), size=(60, 30))
        exitbutton = wx.Button(panel, label="Exit", pos=(10, 100), size=(60, 30))
        stopbutton = wx.Button(panel, label="Stop at end of run", pos=(90, 50), size=(140, 30))
        resetcounterbutton = wx.Button(panel, label="Reset", pos=(90, 100), size=(110, 30))
        runslabel = wx.StaticText(panel, -1, label="Number of runs:", pos=(10, 10))
        self.runstextbox = wx.TextCtrl(panel, -1, "", pos=(100, 5), size=(60, -1))
        self.statusbar = self.CreateStatusBar()
        self.Bind(wx.EVT_BUTTON, self.exit, exitbutton)
        self.Bind(wx.EVT_BUTTON, self.stopWorker, stopbutton)
        self.Bind(wx.EVT_BUTTON, self.startWorker, startbutton)
        self.Bind(wx.EVT_CLOSE, self.closewindow)
        self.bot = DiabloBot()
        self.Bind(wx.EVT_BUTTON, self.resetBot, resetcounterbutton)
        self.bot.addObserver(self)
        self.updateStatusBar()
        self.worker = None

    def resetBot(self, event):
        self.bot.resetCounters()

    def update(self, event):
        self.updateStatusBar()

    def updateStatusBar(self):
        self.statusbar.SetStatusText("legendaries: " + str(self.bot.legendarycount) +
                                    " - " +
                                    "sets: " + str(self.bot.setcount) +
                                    " - " +
                                    "rares: " + str(self.bot.rarecount) +
                                    " - " +
                                    "current run: " + str(self.bot.runcount))

    def startWorker(self, event):
        if not self.worker:
            try:
                worker = WorkerThread(self.bot, int(self.runstextbox.GetValue()))
            except:
                print "could not parse string to integer"
    def exit(self, event):
        self.Close(True)

    def stopWorker(self, event):
        self.worker.stop()
        self.worker = None

    def closewindow(self, event):
        self.Destroy()

if __name__ == '__main__':
    app = wx.PySimpleApp()
    frame = BotGUI(parent=None, id=-1)
    frame.Show()
    app.MainLoop()
