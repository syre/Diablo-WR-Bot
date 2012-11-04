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
import wx.py
import threading

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
        self.pickupLegendary = False
        self.pickupSet = False
        self.pickupRare = False

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
                  'rightofself_click': (650, 315),
                  'last_inventoryspace_coord': (975, 550),
                  'targetleftofself_click': (560, 315),
                  'chest_click': (595, 200),
                  'last_tab_cord': (300, 500),
                  'tp_click': (500, 325),
                  'repairsymbol': (785, 45),
                  'repair_tab': (347, 272),
                  'repair_all': (214, 421),
                  'first_tab': (350, 180),
                  'second_tab': (350, 260),
                  'third_tab': (350, 350),
                  'revive_button': (550,600)}

    gamecolors = {'inventory_empty_block': (19, 11, 5),
                  'rare_itemcolor': (255, 255, 0),
                  'legendary_itemcolor': (191, 100, 47),
                  'set_itemcolor': (0, 255, 0),
                  'red_repaircolor': (216, 0, 0)}

    deathmarkers = [(342, 246), (543, 262), (780, 230)]

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
        self.shell.SendKeys("2")
        self.shell.SendKeys("3")
        self.shell.SendKeys("4")

    def revive(self):
        self.mousePos(self.gamecoords['revive_button'])
        self.leftClick()

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

    def pickupSpecialItems(self, type):
        image = self.screenGrab()
        # y from 100 to 75% of max is the area where items are dropping
        for x in range(1, self.x_1 - self.x_2):
            for y in range(100, self.y_1 / 100 * 75):
                if (image.getpixel((x, y)) == self.gamecolors[type + '_itemcolor']):
                    if (type != 'rare'):
                        self.screenGrabToFile(type)
                    if (self.isDead() and (type != 'rare')):
                        self.revive()
                        self.runFromEntryPoint()
                        for m in range(1, self.x_1 - self.x_2):
                            for n in range(100, self.y_1 / 100 * 75):
                                if (image.getpixel((m, n)) == self.gamecolors[type + '_itemcolor']):
                                    x = m
                                    y = n
                    print type + " item found!"
                    time.sleep(1.5)
                    self.mousePos((x, y))
                    self.leftClick()
                    time.sleep(1.5)
                    if type == 'legendary':
                        self.legendarycount += 1
                        self.pickupSpecialItems(type)
                    elif type == 'set':
                        self.setcount += 1
                        self.pickupSpecialItems(type)
                    elif type == 'rare':
                        self.rarecount += 1
                    self.notifyObservers(type + "count")
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

    def castFrostNova(self):
        self.shell.SendKeys("2")

    def castDiamondSkin(self):
        self.shell.SendKeys("1")

    def castWhirlwind(self):
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
        for mark in self.deathmarkers:
            if (image.getpixel(mark) != (255, 255, 255)):
                return False
        return True

    def moveToNextTab(self):
        image = self.screenGrab()
        if (image.getpixel(self.gamecoords['first_tab']) == (67, 37, 16)):
            self.mousePos(self.gamecoords['second_tab'])
            time.sleep(0.5)
            self.leftClick()
        elif (image.getpixel(self.gamecoords['second_tab']) == (96, 59, 32)):
            self.mousePos(self.gamecoords['third_tab'])
            time.sleep(0.5)
            self.leftClick()
        elif (image.getpixel(self.gamecoords['third_tab']) == (42, 33, 30) and
            image.getpixel(self.gamecoords['last_inventoryspace_coord']) != self.gamecolors['inventory_empty_block']):
            print "no room left"
            self.sleepComputer()

    def sleepComputer(self):
        os.system(r'%windir%\system32\rundll32.exe powrprof.dll,SetSuspendState Hibernate')

    def resetCounters(self):
        self.legendarycount = 0
        self.setcount = 0
        self.rarecount = 0
        self.runcount = 0
        self.notifyObservers("all")

    def enableItemText(self):
        self.shell.SendKeys("-")

    def runFromEntryPoint(self, sleepduration=6.4):
        self.buff()
        self.mousePos((630, 350))
        self.leftDown()
        time.sleep(sleepduration)
        self.leftUp()

    def runArchonBuild(self):
        # archon on rightclick
        self.runFromEntryPoint()
        # cast hydra to the right of self
        self.mousePos(self.gamecoords['rightofself_click'])
        time.sleep(0.5)
        self.rightClick()
        time.sleep(0.5)
        self.castDiamondSkin()
        self.rightDown()
        for i in range(1, 25):
            self.mousePos((800, 280))
            time.sleep(0.5)
            self.mousePos((800, 350))
            time.sleep(0.5)
            self.mousePos((800, 480))
            if (i == 8):
                self.mousePos((150, 300))
                time.sleep(0.1)
                self.leftDown()
                time.sleep(0.5)
                self.leftUp()
            self.shell.SendKeys("1")
        self.leftDown()
        self.rightUp()
        self.castDiamondSkin()

    def castExplosion(self):
        self.rightClick()

    def runCMBuild(self):
        # critical mass / whirlwind: diamond skin, frost nova, teleport, energy armor, whirlwind, explosion
        self.runFromEntryPoint(6.8)
        self.mousePos(self.gamecoords['rightofself_click'])
        time.sleep(0.25)
        for i in range (1, 15):
            time.sleep(0.4)
            self.castWhirlwind()
            if (i > 2):
                self.castFrostNova()
                self.castDiamondSkin()
                self.castExplosion()


    def runGame(self):
        time.sleep(2)
        # refresh diablo window coordinates
        self.refreshDiabloWindow()
        self.resumeDiablo()
        time.sleep(5.5)
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
        # insert build
        self.runArchonBuild()
        self.mousePos(self.gamecoords['rightofself_click'])
        self.leftClick()
        time.sleep(2.6)
        if (self.pickupLegendary):
            self.pickupSpecialItems('legendary')
        if (self.pickupSet):
            self.pickupSpecialItems('set')
        if (self.pickupRare):
            self.pickupSpecialItems('rare')
            self.pickupSpecialItems('rare')
            self.pickupSpecialItems('rare')
        self.exitDiablo()
        time.sleep(10)


class WorkerThread(threading.Thread):
    def __init__(self, function, *args):
        threading.Thread.__init__(self)
        self.function = function
        self.args = args
        self.stop_event = threading.Event()

    def run(self):
        self.function(*self.args)

    def stop(self):
        self.stop_event.set()

    def stopEventSet(self):
        return self.stop_event.isSet()

class BotGUI(wx.Frame, Observer): 
    def __init__(self, parent, id):
        wx.Frame.__init__(self, parent, id, 'PandemoniumBot', size=(300, 200))
        Observer.__init__(self, 'BotGUI')
        panel = wx.Panel(self)
        startbutton = wx.Button(panel, label="Run game", pos=(10, 50), size=(60, 30))
        exitbutton = wx.Button(panel, label="Exit", pos=(10, 100), size=(60, 30))
        stopbutton = wx.Button(panel, label="Stop at end of run", pos=(90, 50), size=(140, 30))
        resetcounterbutton = wx.Button(panel, label="Reset", pos=(90, 100), size=(110, 30))
        self.rarecheckbox = wx.CheckBox(panel, label="rare", pos=(210, 90))
        self.legendarycheckbox = wx.CheckBox(panel, label="legendary", pos=(210, 110))
        self.setcheckbox = wx.CheckBox(panel, label="set", pos=(210, 130))
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

    def runGames(self):
        try:
            numberofruns = int(self.runstextbox.GetValue())
        except:
            print "could not parse string to integer"
            return
        if self.rarecheckbox.IsChecked():
            self.bot.pickupRare = True
        else:
            self.bot.pickupRare = False

        if self.legendarycheckbox.IsChecked():
            self.bot.pickupLegendary = True
        else:
            self.bot.pickupLegendary = False

        if self.setcheckbox.IsChecked():
            self.bot.pickupSet = True
        else:
            self.bot.pickupSet = False

        for i in range(numberofruns):
            self.bot.runGame()
            self.bot.runcount = i + 1
            self.bot.notifyObservers("runcount")
            if self.worker.stopEventSet():
                print "stopping worker"
                return

    def startWorker(self, event):
        if not self.worker or not self.worker.isAlive():
            self.worker = WorkerThread(self.runGames)
            self.worker.start()

    def exit(self, event):
        self.Close(True)

    def stopWorker(self, event):
        self.worker.stop()

    def closewindow(self, event):
        self.Destroy()


def main():
    app = wx.PySimpleApp()
    frame = BotGUI(parent=None, id=-1)
    frame.Show()
    app.MainLoop()

if __name__ == '__main__':
    main()
