#!/usr/bin/env python3
import time
import pygsheets
import csv
import sys
import httplib2
import json
import requests
import googleapiclient.discovery
import googleapiclient.http
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import serial
import argparse
import traceback
import webbrowser

START = b'\x02'
END =   b'\x03'

KEYS = ['A', 'E', 'P']
FORM_URL = "https://docs.google.com/forms/d/e/{id}/formResponse?{data}"

names = {}
emails = {}
CONFIG = {}

DIR=''

def readConfig(dir=DIR):
	global names
	global emails
	global CONFIG
	global DIR

	DIR = dir
	print("I am attendance and the dir is", dir)

	try:
		with open(dir+'people.csv', mode='r') as file:
			reader = csv.reader(file)
			names = dict(reader)
	except FileNotFoundError:
		with open(dir+'people.csv', mode='w+') as file:
			file.close();
	try:
		with open(dir+'ids.csv', mode='r') as file:
			reader = csv.reader(file)
			emails = dict(reader)
	except FileNotFoundError:
		with open(dir+'ids.csv', mode='w+') as file:
			file.close();

	try:
		with open(dir+'attendance.json', 'r') as f:
			CONFIG = json.load(f)
	except FileNotFoundError:
		with open(dir+'attendance.json', mode='w+') as file:
			file.close();

	print("Read config from " + dir)

class MyParser(argparse.ArgumentParser):
	def error(self, message):
		print("error: " + message)
		self.print_help()
		sys.exit(2)

class scanner:
	def __init__(self, port, debounce_delay=.2, GUI=None):
		self.GUI = GUI
		self.debounce_delay=debounce_delay
		self.s = serial.Serial(port, timeout=debounce_delay)

	def listPorts(self):
		return list(serial.Serial.list_ports.comports())

	def connect(self, port):
		self.s = serial.Serial(port, timeout=self.debounce_delay)

	def readID(self):
		time.sleep(5)
		id = b''
		while(True):
			if(self.GUI is not None and self.GUI.killTrigger.isSet()):
				raise KeyboardInterrupt
			c = self.s.read()
			if(c == START):
				id = b''
			elif(c == END):
				id = id.decode("utf-8")
				if(self.GUI is not None):
					self.GUI.playSound()
				log(id,GUI=self.GUI)
				break
			else:
				id += c
		while(len(self.s.read()) > 0): #Don't Return ID until tag is removed
			pass
		return id

def round(number, nearest):
	return number // nearest * nearest

def fillRow(sheet, row, data, headers=None):
	for col in range(1,sheet.col_count+1):
		key = sheet.cell(1,col).value
		try:
			sheet.update_cell(row, col, data[key])
		except KeyError:
			sheet.update_cell(row, col, "")

def submitForm(form_id, data):
	data_string = ''
	for i,keyvalue in enumerate(data.items()):
		key,value = keyvalue
		data_string += "{}={}".format(key,value)
		if(i+1 != len(data)):
			data_string += '&'
	url = FORM_URL.format(id=form_id, data=data_string)
	print(url)
	try:
		status = requests.get(url, timeout=1, headers={'Connection':'close'}).status_code
	except requests.exceptions.ConnectionError:
		status = 408
	#print(response)
	return status

def markFromGUI(person, state, date, GUI=None):
	try:
		mark(person, state, date, GUI=GUI)
	except Exception as e:
		log("An exception has occured: {}".format(str(e)), GUI=GUI)
		traceback.print_exc()
		raise

def processFromGUI(GUI):
	try:
		processv2(GUI)
	except Exception as e:
		log("An exception has occured: {}".format(str(e)), GUI=GUI)
		traceback.print_exc()
		raise

def mark(person, state, date, GUI=None):
	data = {
		CONFIG["EMAIL"]: person,
		CONFIG["STATE"]: state,
		CONFIG["DATE"]:  "{year}-{month:0{width}}-{day:0{width}}".format(year=date.year, month=date.month, day=date.day, width=2),
	}

	status = submitForm(CONFIG["FORM_ID"], data)
#	print(status)
	if(status == 400):
		log("Error 400: Bad Request.\nFix invalid input. (Check that email is valid)", GUI=GUI)
	elif(status != 200):
		log("Error "+str(status), GUI=GUI)
		log("Could not connect to google. Saving offline.", GUI=GUI)
		with open(DIR+"offline.csv", 'w+') as file:
			file.write("{},{},{}".format(person, state, date))
	else:
		try:
			name = names[person]
		except KeyError:
			name = person.lower()
		if(state=='P'):
			log("Welcome, {}\n".format(name), GUI=GUI)
		elif(state=='E'):
			log("Marked {} as excused\n".format(name), GUI=GUI)
		elif(state=='A'):
			log("Marked {} as absent\n".format(name), GUI=GUI)


def log(msg, end="\n", GUI=None):
	print(msg, end=end)
	if(GUI is not None):
		GUI.log(msg, end=end)

def takeAttendanceFromGUI(GUI, port):
	try:
		takeAttendance(port, GUI)
	except KeyboardInterrupt:
		log("Stopped.", GUI=GUI)
	except Exception as e:
		log("An exception has occured: {}".format(str(e)), GUI=GUI)
		GUI.switchToStart()
		raise
#	finally:
#		GUI.resetWhateverHasToBeReset()

def takeAttendance(port, GUI=None):
	global emails

	log("Connecting to serial... ", end='', GUI=GUI)
	sys.stdout.flush()
	scan = scanner(port, GUI=GUI)
	log("Connected.\nReady to take attendance.", GUI=GUI)

	while(True):
		log("ID:", end=' ', GUI=GUI)
		sys.stdout.flush()
		id = scan.readID()
#		log(id, GUI=GUI)

		isMember = True
		if(id not in emails.keys()):
			isMember = register(id, GUI)
		if(isMember):
			email = emails[id]
			try:
				name = names[email]
			except KeyError:
				name = email.lower()
			log("Welcome, {}\n".format(name), GUI=GUI)
			mark(email, 'P', datetime.date.today())

def manual(state='P'):
	while(True):
		valid = False
		while(not valid):
			email = input("Email: ").lower()
			valid = (email in (key.lower() for key in names.keys()))
			if(not valid):
				print("Invalid school email.")
		mark(email, state, datetime.date.today())

def updateOldEntries():
	while(True):
		valid = False
		while(not valid):
			email = input("Email: ")
			valid = (email.lower() in (key.lower() for key in names.keys()))
			if(not valid):
				print("Invalid school email.")
		state = input("State: ").upper()
		date_string = input("Date: ")
		month, day, year = [ int(s) for s in date_string.split('/') ]
		date = datetime.date(year, month, day)
		mark(email, state, date)

def register(id, GUI=None):
	try:
		valid = False
		while(not valid):
			if(GUI is not None):
				valueHolder = GUI.getEmail()
				while(valueHolder.getValue() is None):
					time.sleep(0.1)
				email = valueHolder.getValue()
				if(email is False):
					raise ValueError
			else:
				email = input("Enter you school email: ")
#			valid = (email.lower() in (key.lower() for key in names.keys()))
			valid = ("@" in email)
			if(not valid):
				if(GUI is not None):
					GUI.complainAbout("Invalid school email.")
				else:
					print("Invalid school email.")
		emails[id] = email
		with open(DIR+'ids.csv', 'a') as f:
			f.write("{},{}\n".format(id,email))
		return True
	except ValueError:
		log("Canceling new member.", GUI=GUI)
		return False

class person():
	def __init__(self, email, dates):
		self.dates = dict((date,'') for date in dates)
		self.email = email

	def mark(self, date, value):
		if(date in self.dates.keys()):
			self.dates[date] = value

	def getRow(self):
		print(self.dates)
		values = list(self.dates.values())
		row = [self.email]
		row.append(values.count("P"))
		row.append(values.count("E"))
		row.append(values.count("A"))
		row.extend(values)
		print(row)
		return row

def processv2(GUI=None):
	log("Reading Responses...", GUI=GUI)

	# Login
	http_client = httplib2.Http(timeout=2)
	log("Authorizing...", end=' ', GUI=GUI)
	try:
		gc = pygsheets.authorize(outh_file=DIR+'client_secret.json', outh_creds_store=DIR, no_cache=True)
#		gc = pygsheets.authorize(outh_file=DIR+'client_secret.json', http_client=http_client, retries=10)
	except:
		log("Could not connect to the server.", GUI=GUI)
		raise
	log("Done.", GUI=GUI)

	#Open sheet from config
	sheet = gc.open_by_key(CONFIG["SHEET_ID"])
	responses = sheet.worksheet_by_title("Responses")
	overview = sheet.worksheet_by_title("Overview")
	log("Opened Spreadsheet.", GUI=GUI)

	#Read data
	data = responses.get_all_records()
	headers = overview.get_row(1)
	meetings = [date for date in  headers[4:] if date != '']  #Dates start in fourth col
	log("Read Data.", GUI=GUI)

	#Build output matrix
	people = {}
	for row in data:
		try:
			email = row["Student Email"].lower()
			if(email not in people.keys()):
				people[email] = person(email, meetings)
			people[email].mark(row["Date"], row["State"])
		except KeyError:
			log(":",str(row), GUI=GUI) #No idea why there would be a key error
	people = people.values()
	people = sorted(people, key=lambda person: person.email)
	print(people)
	output_matrix = [person.getRow() for person in people]

	overview.resize(rows=len(output_matrix)+1, cols=len(output_matrix[0]))
	overview.update_cells('A2:{}'.format(rowcol_to_a1(len(output_matrix)+1,len(output_matrix[0]))), output_matrix)
	log("Wrote Data. Done!", GUI=GUI)


def process(GUI=None):
	log("Reading Responses...", GUI=GUI)
	# Login to  google sheets
# 	scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
# 	http_client = httplib2.Http(timeout=2)
# 	client = pygsheets.authorize(outh_file=DIR+"client_secret.json", http_client=http_client, retries=10)

	http_client = httplib2.Http(timeout=2)
	log("Authorizing...", end=' ', GUI=GUI)
	try:
		gc = pygsheets.authorize(client_secret=DIR+'client_secret.json', outh_file=DIR+'client_secret.json', http_client=http_client, retries=10)
	except:
		log("Could not connect to the server.", GUI=GUI)
		raise
	log("Done.", GUI=GUI)


	# Find a Sheet by name and open the first sheet
	sheet = gc.open_by_key(CONFIG["SHEET_ID"])
	responses = sheet.worksheet_by_title("Responses")
	overview = sheet.worksheet_by_title("Overview")

# 	responses = spreadsheet.worksheet("Responses")
# 	overview = spreadsheet.worksheet("Overview")

	#Put the read functions first and together because they take a long time.
	data = responses.get_all_records()
	headers = overview.get_row(1)
	meetings = dict((date, '') for date in headers[4:])
# 	meetings = dict((date, '') for date in overview.row_values(1)[1:])

	log("Processing Data...", GUI=GUI)

	#First get the meetings and have the ones in the past default to Absent
	today = datetime.date.today()
	for meeting in meetings:
		try:
			month, day, year = [ int(s) for s in meeting.split('/') ]
			meeting_date = datetime.date(year, month, day)
			if(meeting_date < today):
				meetings[meeting] = 'A'
		except ValueError: #Not a meeting, some other header
			pass

	# Create an dict for each person and fill it's data
	people = {}

	for row in data:
		email = row["Student Email"]
		date = row["Date"]
		if(email not in people):
			people[email] = meetings.copy()
			try:
				people[email]["Name"] = names[email]
			except KeyError:
				people[email]["Name"] = email
		if(date in meetings):
			people[email][date] = row["State"]

	log("Resizing Spreadsheet...", GUI=GUI)

	#Resize the spreadsheet
# 	overview.resize(rows=2)
	if(len(people) > 0):
		pass
# 		overview.resize(rows=len(people)+1)
	else:
# 		overview.resize(rows=2)
		fillRow(overview, 2, {})

	log("Processed data. Updating spreadsheet...", GUI=GUI)

	last_cell = rowcol_to_a1(overview.row_count,overview.col_count)
	cell_list = overview.range('A2:{}'.format(last_cell))

	rows = [cell_list[x:x+overview.col_count] for x in range(0, len(cell_list), overview.col_count)]

	FIRST_ROW = overview.row_values(1)

	# Update the sheet
	people = sorted(people.values(), key=lambda k: k['Name'])
	for r,person in enumerate(people):
		for key in KEYS:
			person[key] = list(person.values()).count(key)
		for c,key in enumerate(FIRST_ROW):
			rows[r][c].value = person[key]
	overview.update_cells(cell_list)

ASCII_START = 64
def rowcol_to_a1(row, col):
	div = col
	column_label = ''
	while div:
		(div, mod) = divmod(div, 26)
		if mod == 0:
			mod = 26
			div -= 1
		column_label = chr(mod + ASCII_START) + column_label

	label = '%s%s' % (column_label, row)

	return label

def openSheet():
	webbrowser.open("https://docs.google.com/spreadsheets/d/{}/".format(CONFIG["SHEET_ID"]))

if(__name__ == "__main__"):
	parser = MyParser(description='A script for mannaging the attendance of an orginization or a class.')
	parser.add_argument("-m", "--manual",  action='store_true', help="Take manual attendance.")
	parser.add_argument("-e", "--excuse",  action='store_true', help="Excuse people from today's meeting.")
	parser.add_argument("-t", "--take",    action='store_true', help="Take attendance normally.")
	parser.add_argument("-u", "--update",  action='store_true', help="Update older entries.")
	parser.add_argument("-p", "--process", action='store_true', help="Process the Spreadsheet overview.")
	args = parser.parse_args()

	readConfig("SLEHS 2018-19/")

	try:
		if(args.update):
			updateOldEntries()
		elif(args.manual):
			manual()
		elif(args.excuse):
			manual(state='E')
		elif(args.take):
			takeAttendance()
	except KeyboardInterrupt:
		print("\n")
	if(args.process):
		process()
	print("Done!")

