#!/usr/bin/env python3
import wx
import wx.adv
import wx.lib.intctrl
import wx.lib.masked.numctrl
import json
import hours
import threading
import webbrowser
import time
import serial
import serial.tools.list_ports
import calendar
import datetime
import attendance
import os
import sys
import traceback

IS_WINDOWS = (os.name == 'nt')

BEEP_PATH = "sfx/beep.wav"

if getattr(sys, 'frozen', False):
	application_path = os.path.dirname(sys.executable)
	if(not IS_WINDOWS):
		BEEP_PATH = "/".join(application_path.split("/")[:-1]) + "/Resources/" + BEEP_PATH
		application_path = "/".join(application_path.split("/")[:-3])
#	else: #Only if not in one file mode on windows
#		application_path = "//".join(application_path.split("\\")[:-1])
#elif __file__:
#	application_path = os.path.dirname(__file__)
#	print("File:",__file__)
else:
	application_path = "."

print("The application path is \"{}\"".format(application_path))

try:
	with open("/Users/alex/Desktop/theDir.txt", "w+") as file:
		print("Parent Dir:",application_path, file=file)
		print("Beep:",BEEP_PATH, file=file)
		file.close()
except FileNotFoundError:
	pass

#with open(application_path + "/foo.txt", "w+") as file:
#	file.write(application_path)
#	file.close()

THREADED_TIME_ESTIMATE = False

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "June", "July", "Aug", "Sept", "Oct", "Nov", "Dec"]
year = datetime.datetime.now().year
YEARS = [str(i) for i in range(year-2,year+6)]
# Define the tab content as classes:
class HoursTab(wx.Panel):
	def __init__(self, parent):
		self.updateThread = None
		self.timeThread = None
		self._parent = parent
		wx.Panel.__init__(self, parent)
		self.updateBtn = wx.Button(self, label="Update New")
		self.updateBtn.Bind(wx.EVT_BUTTON,self.update)
		self.forceUpdateBtn = wx.Button(self, label="Force Update Everything")
		self.forceUpdateBtn.Bind(wx.EVT_BUTTON,self.forceUpdate)
		self.viewBtn = wx.Button(self, label="Open Overview")
		self.viewBtn.Bind(wx.EVT_BUTTON,self.viewSheet)

		sizer2 = wx.BoxSizer(wx.HORIZONTAL)
		sizer2.Add(self.updateBtn, 1, wx.EXPAND)
		sizer2.Add(self.forceUpdateBtn, 1, wx.EXPAND)
#		sizer2.Add(self.viewBtn, 1, wx.ALIGN_RIGHT)

		self.console = wx.TextCtrl(self, style=wx.TE_READONLY|wx.TE_MULTILINE)
		self.console.SetEditable(False)

		self.gauge = wx.Gauge(self, range=100, style=wx.GA_HORIZONTAL)
		self.gaugeTime = wx.StaticText(self, label="", style=wx.ALIGN_LEFT)
		self.gaugeText = wx.StaticText(self, label="", style=wx.ALIGN_RIGHT|wx.ST_NO_AUTORESIZE)

		sizer3 = wx.BoxSizer(wx.HORIZONTAL)
		sizer3.Add(self.gaugeTime, 0, wx.EXPAND,wx.LEFT)
		sizer3.Add(self.gaugeText, 1, wx.RIGHT)

		if(IS_WINDOWS): #For some reason windows likes to cut the progress bar off and the right edge of the text
			sizer3.Add((10,43), 0, wx.RIGHT)

		sizer1 = wx.BoxSizer(wx.VERTICAL)
		sizer1.Add(sizer2, 0, wx.TOP|wx.CENTER)
		sizer1.Add((-1,5))
		sizer1.Add(self.console, 1, wx.EXPAND)
		sizer1.Add(self.gauge, 0, wx.EXPAND)
		sizer1.Add(sizer3, 0, wx.EXPAND)
		parent.SetSizer(sizer1)

		self.gaugeText.SetLabel("Ready")

	def update(self,commandevent):
		if(self.updateThread is None or not self.updateThread.isAlive()):
			self.gaugeText.SetLabel("Connecting...")
			print("Running new thread")
			self.updateThread = threading.Thread(target = hours.updateFromGUI, args = (self,False))
			self.updateThread.daemon = True
			self.updateThread.start()
		else:
			msg = "Already updating hours."
			wx.MessageBox(msg, "Error")
			print(msg)

	def forceUpdate(self,commandevent):
		if(self.updateThread is None or not self.updateThread.isAlive()):
			self.gaugeText.SetLabel("Starting")
			print("Running new thread")
			self.updateThread = threading.Thread(target = hours.updateFromGUI, args = (self,True))
			self.updateThread.daemon = True
			self.updateThread.start()
		else:
			msg = "Already updating hours."
			wx.MessageBox(msg, "Error")
			print(msg)

	def viewSheet(self, event):
		hours.openSheet()

	def cancel(self,commandevent):
		if(self.updateThread is not None):
			self.updateThread = None

	def log(self, msg, end=" "):
		wx.CallAfter(self.console.AppendText, msg + end)

	def pulseProgress(self):
		wx.CallAfter(self.gauge.Pulse)

	def setProgressRange(self, range):
		self.range = range
		wx.CallAfter(self.gauge.SetRange, (range))

	def countDown(self):
		wx.CallAfter(self._updateTimeRemaining,self.time)
		while(self.time is not None and self.updateThread is not None):
			if(self.updateThread is None):
				print("Update Thread is None")
			time.sleep(1.1)
			if(self.time > 0 and self.time > self.lastTime-self.timeDelta):
				self.time -= 1
			wx.CallAfter(self._updateTimeRemaining,self.time)
		print("time is None or update thread is None")
		self.updateThread = None
		self.gaugeText.SetLabel("")
		wx.CallAfter(self._updateTimeRemaining,None)

	def setProgress(self, progress, time=None, timeDelta=1):
		if(THREADED_TIME_ESTIMATE):
			self.time = time
			self.lastTime = time
			self.timeDelta = timeDelta
			if(self.timeThread is None and self.time is not None):
				self.timeThread = threading.Thread(target=self.countDown)
				self.timeThread.start()
			if(self.time is None):
				self.timeThread = None
		wx.CallAfter(self.console.AppendText, "Done!\n")
		wx.CallAfter(self._setProgress, progress,time)

	def clearProgressBar(self, msg=""):
		wx.CallAfter(self._clearProgressBar, msg)

#	def testTimeRemaining(self, initialValue=7500, timedelta=1):
#		while(initialValue >= 0):
#			wx.CallAfter(self._updateTimeRemaining, (initialValue))
#			time.sleep(timedelta)
#			initialValue -= 1

	def _setProgress(self, progress, time):
		self.gauge.SetValue(progress)
		self.gaugeText.SetLabel(str(progress)+"/"+str(self.range))

		if((time is not None and not THREADED_TIME_ESTIMATE) or True):
			self._updateTimeRemaining(time)

	def _clearProgressBar(self, msg):
		self.gauge.SetValue(0)
		self.gaugeText.SetLabel(msg)
		self.gaugeTime.SetLabel("")

	def _updateTimeRemaining(self, s):
		m, s = divmod(s, 60)
		h, m = divmod(m, 60)
		timestr = ""
		if(h == 1):
			timestr = "1 hour "
		elif(h > 0):
			timestr = "{} hours ".format(h)

		if(m == 1):
			timestr += "1 minute "
		elif(m > 0):
			timestr += "{} minutes ".format(m)

		if(s == 1):
			timestr += "1 second "
		elif(s > 0):
			timestr += "{} seconds ".format(s)
		elif(timestr == ""):
			timestr = "no time "

		self.gaugeTime.SetLabel("About {}remaining".format(timestr))

class AttendanceTab(wx.Panel):
	def __init__(self, parent):
		wx.Panel.__init__(self, parent)
		self.updateThread = None

		self.console = wx.TextCtrl(self, style=wx.TE_READONLY|wx.TE_MULTILINE)
		self.console.SetEditable(False)

		nb = wx.Notebook(self)

		self.tab1 = TakeAttendanceTab(nb, self.console)
		self.tab2 = ManualAttendanceTab(nb, self.console, state='P')
		self.tab3 = ManualAttendanceTab(nb, self.console, state='E')

		nb.AddPage(self.tab1, "Scan In")
		nb.AddPage(self.tab2, "Manual")
		nb.AddPage(self.tab3, "Excuse")

		self.process = wx.Button(self, label="Process Spreadsheet")
		self.view = wx.Button(self, label="View Spreadsheet")

		self.process.Bind(wx.EVT_BUTTON, self.processSheet)
		self.view.Bind(wx.EVT_BUTTON, self.viewSheet)


		sizer2 = wx.BoxSizer(wx.HORIZONTAL)
		sizer2.Add(self.process, 0, wx.CENTER)
		sizer2.Add(self.view, 0, wx.CENTER)

		sizer1 = wx.BoxSizer(wx.VERTICAL)
		sizer1.Add(nb, 1, wx.EXPAND)
		sizer1.Add(sizer2, 0, wx.CENTER)
		sizer1.Add((-1,5))

		sizer = wx.BoxSizer(wx.HORIZONTAL)
		sizer.Add(self.console, 1, wx.EXPAND)
		sizer.Add(sizer1, 1, wx.EXPAND)
		self.SetSizer(sizer)

	def processSheet(self, event=None):
		if(self.updateThread is None or not self.updateThread.isAlive()):
			print("Processing Attendance Sheet")
			self.updateThread = threading.Thread(target = attendance.processFromGUI, args = (self,))
			self.updateThread.daemon = True
			self.updateThread.start()
		else:
			msg = "Already processing spreadsheet."
			wx.MessageBox(msg, "Error")
			print(msg)

	def viewSheet(self, event):
		attendance.openSheet()

	def log(self, msg, end=" "):
		wx.CallAfter(self.console.AppendText, msg + end)

class TakeAttendanceTab(wx.Panel):
	def __init__(self, parent, console, id=None):
		wx.Panel.__init__(self, parent)
		self.updateThread = None
		self.killTrigger = threading.Event()
		self.console = console
		self.refresh = wx.Button(self, label="Refresh")
		self.refresh.Bind(wx.EVT_BUTTON, self.refreshPorts)
		self.portChooser = wx.Choice(self)
		self.portChooser.Bind(wx.EVT_CHOICE, self.OnChoice)

		sizer = wx.BoxSizer(wx.HORIZONTAL)
		sizer.Add(self.portChooser, 1, wx.EXPAND)
		sizer.Add(self.refresh, 0, wx.RIGHT)

		self.startBtn = wx.Button(self, label="Start")
		self.startBtn.Bind(wx.EVT_BUTTON, self.takeAttendance)

		sizer2 = wx.BoxSizer(wx.VERTICAL)
		sizer2.Add(sizer,0,wx.TOP|wx.EXPAND)
		sizer2.Add((-1,5), 0)
		sizer2.Add(self.startBtn, 0, wx.CENTER)
		self.SetSizer(sizer2)

	def switchToStart(self):
		self.startBtn.SetLabel("Start")
		self.startBtn.Bind(wx.EVT_BUTTON, self.takeAttendance)

	def switchToStop(self):
		self.startBtn.SetLabel("Stop")
		self.startBtn.Bind(wx.EVT_BUTTON, self.killThread)

	def killThread(self, event=None):
		self.killTrigger.set()
		self.updateThread.join()
		self.switchToStart()

	def getEmail(self):
		value = ValueObject()
		wx.CallAfter(NewMember, self, value)
		return value

	def complainAbout(self, message):
		dlg = wx.MessageDialog(self, message, "Error", wx.OK | wx.ICON_WARNING)
		dlg.ShowModal()
		dlg.Destroy()

	def refreshPorts(self, event=None):
		ports = serial.tools.list_ports.comports()
		portStrings = []
		self.realPorts = []
		for port in ports:
			if(port[1] != "n/a"):
				portStrings.append(str(port[0])+": "+str(port[1]))
				self.realPorts.append(port[0])
		self.portChooser.Clear()
		self.portChooser.AppendItems(portStrings)
		self.portChooser.SetSelection(0)
#		if(len(portStrings) == 0):
#			self.portChooser.Disable()
#			self.startBtn.Disable()
#		else:
#			self.portChooser.Enable()
#			self.startBtn.Enable()

	def takeAttendance(self,commandevent):
		if(self.updateThread is None or not self.updateThread.isAlive()):
			try:
				port = self.realPorts[self.portChooser.GetSelection()]
				print("Running new attendance thread")
				self.killTrigger.clear()
				self.updateThread = threading.Thread(target = attendance.takeAttendanceFromGUI, args = (self,port))
				self.updateThread.daemon = True
				self.updateThread.start()
				self.switchToStop();
			except IndexError:
				msg = "Must connect a scanner first."
				wx.MessageBox(msg, "Error")
				print(msg)
		else:
			msg = "Already taking attendance."
			wx.MessageBox(msg, "Error")
			print(msg)

	def log(self, msg, end=" "):
		wx.CallAfter(self.console.AppendText, msg + end)

	def OnChoice(self, event):
		print(event)

	def playSound(self):
		self.sound = wx.adv.Sound(BEEP_PATH)
		if self.sound.IsOk():
			self.sound.Play(wx.adv.SOUND_ASYNC)

class ValueObject():
	def __init__(self):
		self.value = None

	def getValue(self):
		return self.value

	def setValue(self, value):
		self.value = value

#Dialog Popup
class NewMember(wx.Dialog):
	def submit(self, event):
		self.valueHolder.setValue(self.emailInput.GetValue())
		self.Destroy()

	def OnClose(self, event):
		self.valueHolder.setValue(False)
		self.Destroy()

	def __init__(self, parent, valueHolder):
		wx.Dialog.__init__(self, parent)#message="Enter an email address for tag #{}".format(id)) #,title="New Tag Registration")
		self.Center()
		self.done = False
		self.valueHolder = valueHolder

		emailLabel = wx.StaticText(self, label='Email ')
		self.emailInput = wx.TextCtrl(self)

		cancelBtn = wx.Button(self, label="Cancel")
		cancelBtn.Bind(wx.EVT_BUTTON, self.OnClose)
		registerBtn = wx.Button(self, label="Register")
		registerBtn.Bind(wx.EVT_BUTTON, self.submit)

		inputSizer = wx.BoxSizer(wx.HORIZONTAL)
		inputSizer.Add((5,-1), 0)
		inputSizer.Add(emailLabel, 0, wx.ALIGN_CENTER_VERTICAL)
		inputSizer.Add(self.emailInput, 1, wx.EXPAND)
		inputSizer.Add((10,-1), 0)

		buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
		buttonSizer.Add((-1,-1), 1)
		buttonSizer.Add(cancelBtn, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL)
		buttonSizer.Add(registerBtn, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL)
		buttonSizer.Add((15,-1), 0)

		sizer = wx.BoxSizer(wx.VERTICAL)
		sizer.Add((-1,-1), 1)
		sizer.Add(inputSizer, 0, wx.EXPAND)
		sizer.Add(buttonSizer, 3, wx.EXPAND)

		self.SetSizer(sizer)

		self.Bind(wx.EVT_CLOSE, self.OnClose)

		self.SetSize((360,100))

		self.Center()
		self.Show()

class ManualAttendanceTab(wx.Panel):
	def __init__(self, parent, console, state='P'):
		wx.Panel.__init__(self, parent)

		self.console = console
		self.state = state

		emailLabel = wx.StaticText(self, label='Email ')

		self.emailInput = wx.TextCtrl(self)
		self.submit = wx.Button(self, label="Submit")
		self.submit.Bind(wx.EVT_BUTTON, self.mark)

		sizer = wx.BoxSizer(wx.HORIZONTAL)
		sizer.Add(emailLabel, 0, wx.ALIGN_CENTER_VERTICAL)
		sizer.Add(self.emailInput, 1, wx.EXPAND)
		sizer.Add(self.submit, 0)

		self.month = wx.Choice(self, choices = MONTHS)
		self.day = wx.Choice(self)
		self.year = wx.Choice(self, choices = YEARS)
		self.month.Bind(wx.EVT_CHOICE, self.fixDays)
		self.year.Bind(wx.EVT_CHOICE, self.fixDays)

		sizer2 = wx.BoxSizer(wx.HORIZONTAL)
		sizer2.Add(self.month,1)
		sizer2.Add(self.day,1)
		sizer2.Add(self.year,1)

		sizer1 = wx.BoxSizer(wx.VERTICAL)
		sizer1.Add(sizer,0,wx.TOP|wx.EXPAND)
		sizer1.Add((-1, 5))
		sizer1.Add(sizer2,0,wx.EXPAND)

		self.SetSizer(sizer1)

		self.setDay(datetime.datetime.now())

	def setDay(self, now):
#		self.year.setSelection(
		self.month.SetSelection(now.month-1)
		self.year.SetSelection(now.year - int(YEARS[0]))

		self.fixDays()
		self.day.SetSelection(now.day-1)

	def fixDays(self, event=None):
		day = self.day.GetSelection()
		days = calendar.monthrange(int(self.year.GetString(self.year.GetSelection())), self.month.GetSelection()+1)[1]
		DAYS = [str(i) for i in range(1,days+1)]
		self.day.Clear()
		self.day.AppendItems(DAYS)
		if(day+1 > days):
			day = days-1
		self.day.SetSelection(day)

	def mark(self,event):
		email = self.emailInput.GetValue()
		date = datetime.date(int(YEARS[self.year.GetSelection()]), self.month.GetSelection()+1, self.day.GetSelection()+1)
#		print("Email:",email,"\nDate:",date,"\nState:",self.state)
		attendance.markFromGUI(email, self.state, date, GUI=self)

	def log(self, msg, end=" "):
		wx.CallAfter(self.console.AppendText, msg + end)

class MainFrame(wx.Frame):
	def __init__(self):
		wx.Frame.__init__(self, None, title="NHS Management Software")

		# Create a panel and notebook (tabs holder)
		p = wx.Panel(self)
		nb = wx.Notebook(p)

		# Year Selection Settings
		yearLabel = wx.StaticText(p, label='Year ', style=wx.ALIGN_CENTER_VERTICAL)

#		blacklist = ['build', 'dist', 'src', '__pycache__',]
#		configs = [x for x in next(os.walk('.'))[1] if x not in blacklist]
		configs = [x for x in next(os.walk('.'))[1] if os.path.isfile(x+"/client_secret.json") ]

		self.chooser = wx.Choice(p, choices = configs)
		self.chooser.Bind(wx.EVT_CHOICE, self.OnChoice)
		self.chooser.SetSelection(0)

		self.newYear = wx.Button(p, label='Add New')
		self.newYear.Bind(wx.EVT_BUTTON, self.AddNew)

		yearSizer = wx.BoxSizer(wx.HORIZONTAL)
		yearSizer.Add((5, -1))
		yearSizer.Add(yearLabel, 0, wx.ALIGN_CENTER_VERTICAL)
		yearSizer.Add(self.chooser, 1, wx.EXPAND)
		yearSizer.Add(self.newYear, 0, wx.EXPAND)
		yearSizer.Add((5, -1))

		# Create the tab windows
		self.tab1 = HoursTab(nb)
		self.tab2 = AttendanceTab(nb)

		# Add the windows to tabs and name them.
		nb.AddPage(self.tab1, "Hours")
		nb.AddPage(self.tab2, "Attendance")

		# Set noteboook in a sizer to create the layout
		sizer = wx.BoxSizer(wx.VERTICAL)
		sizer.Add((-1, 5))
		sizer.Add(yearSizer, 0, wx.EXPAND|wx.LEFT|wx.RIGHT|wx.TOP)
		sizer.Add(nb, 1, wx.EXPAND)
		p.SetSizer(sizer)
		self.SetMinSize((640,360))
		self.SetSize((640,360))

		self.tab2.tab1.refreshPorts()
		self.OnChoice(None)
#		self.Center()

#		self.Bind(wx.EVT_CLOSE, self.KillThreads)

	def log(self, msg, end=" "):
		wx.CallAfter(self.tab1.console.AppendText, msg + end)

	def OnChoice(self, event):
		configFolder = self.chooser.GetString(self.chooser.GetSelection())
		print("Selected \""+ configFolder +"\" from Combobox")
		attendance.readConfig(dir=application_path+"/"+configFolder+"/")
		hours.readConfig(dir=application_path+"/"+configFolder+"/", GUI=self)

	def AddNew(self, event):
		dlg = wx.MessageDialog(self, "Not yet implemented", "Error", wx.OK | wx.ICON_WARNING)
		dlg.ShowModal()
		dlg.Destroy()

if __name__ == "__main__":
	app = wx.App()
	MainFrame().Show()
	app.MainLoop()
