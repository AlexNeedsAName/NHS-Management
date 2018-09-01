#!/usr/bin/env python3
from collections import namedtuple
import csv
import datetime
import httplib2
import json
import time
import pygsheets
import sys
import socket
import traceback
import webbrowser

Entry = namedtuple('Entry', ['date', 'task', 'hours', 'contact', 'photo'])

REQUIRED_IN_HOURS = 10
REQUIRED_OUT_HOURS = 10

INDIVIDUAL_SHEETS_NAME = "{name}'s {year} Hours"

LINK = "=HYPERLINK(\"https://docs.google.com/spreadsheets/d/{id}\", \"Click for Detail\")"

TIME = 0
EMAIL = 1
DATE = 2
TASK = 3
NUM = 4
TYPE = 5
CONTACT = 6
PHOTO = 7

NONE= 0
LOW = 1
MED = 2
HIGH= 3
ALL = 4

LOG_LEVEL = ALL

CHANGED_DELAY = 5
UNCHANGED_DELAY = 1

def log(msg, level=HIGH, end='\n', flush=True, GUI=None):
	if(level <= LOG_LEVEL):
		print(msg, end=end, flush=flush)
	if(GUI is not None):
		GUI.log(msg, end=end)
class Person:
	def __init__(self, email, name):
		self.email = email
		self.name = name
		self.in_hours = Hours(REQUIRED_IN_HOURS)
		self.out_hours = Hours(REQUIRED_OUT_HOURS)
		self.sheet_id = None
		self.last_check = None

	def addLastCheck(self, last_check):
		self.last_check = last_check

	def addSheet(self, sheet_id):
		self.sheet_id = sheet_id

	def addHours(self, row):
		if(row[TYPE] == 'In Hours'):
			hours = self.in_hours
		else:
			hours = self.out_hours
		hours.addEntry(row)

	def getTotal(self):
		return self.in_hours.getTotal() + self.out_hours.getTotal()

	def getRemaining(self):
		return self.in_hours.getRemaining() + self.out_hours.getRemaining()

	def getOverview(self):
		return [self.email,
		        self.name,
		        self.in_hours.getTotal(),
		        self.in_hours.getRemaining(),
		        self.out_hours.getTotal(),
		        self.out_hours.getRemaining(),
		        LINK.format(id=self.sheet_id),
		]

	def getPersonalOverview(self):
		return [self.name,
		        self.in_hours.getTotal(),
		        self.in_hours.getRemaining(),
		        self.out_hours.getTotal(),
		        self.out_hours.getRemaining(),
		        self.last_check
		]

class Hours:
	def __init__(self, required_hours):
		self.total = 0
		self.required = required_hours
		self.entries = []

	def addEntry(self, row):
		entry = Entry(
			row[DATE],
			row[TASK],
			float(row[NUM]),
			row[CONTACT],
			row[PHOTO],
		)
		self.entries.append(entry)
		self.total += entry.hours

	def getEntries(self):
		return self.entries

	def getTotal(self):
		return self.total

	def getRemaining(self):
		remaining = self.required - self.total
		if(remaining < 0):
			remaining = 0
		return remaining

	def getMatrix(self):
		if(len(self.entries) != 0):
			matrix = [list(entry) for entry in self.entries]
		else:
			 matrix = [['']]
		return matrix

DIR=''

def readConfig(dir=DIR, GUI=None):
	global CONFIG
	global names
	global DIR

	DIR = dir

	with open(dir+"hours.json", 'r') as f:
		CONFIG = json.load(f)

	#Open the csv of emails to get people's real names.
	with open(dir+'people.csv', mode='r') as file:
		reader = csv.reader(file)
		names = dict(reader)

	print("Read config from " + dir)

def writeConfig(CONFIG, dir=DIR):
	with open(dir+"hours.json", 'w+') as f:
		f.write(json.dumps(CONFIG, indent=4))

def updateWorksheet(worksheet, matrix):
	worksheet.resize(rows=len(matrix)+1, cols=5)
	worksheet.update_cells('A2:E{}'.format(len(matrix)+1), matrix)

def updateFromGUI(GUI, force=False):
	try:
		updateHours(GUI, force)
	except Exception as e:
		log("An exception has occured: {}".format(str(e)), end="\n", flush=True, level=HIGH, GUI=GUI)
		traceback.print_exc()
		GUI.clearProgressBar()
		raise
	finally:
		GUI.clearProgressBar("Ready")

def updateHours(GUI=None, force=False):
	if(GUI is not None):
		GUI.pulseProgress()

# 	CONFIG = readConfig()

	#Connect to google and open the sheet.
	http_client = httplib2.Http(timeout=2)
	log("Authorizing...", end=' ', flush=True, level=LOW, GUI=GUI)
	try:
#		gc = pygsheets.authorize(outh_file=DIR+"client_secret.json", outh_creds_store=DIR, http_client=http_client, retries=10, no_cache=True)
		gc = pygsheets.authorize(outh_file=DIR+'client_secret.json', outh_creds_store=DIR, no_cache=True, http_client=http_client, retries=10)
	except:
		log("Could not connect to the server.", end="\n", flush=True, level=HIGH, GUI=GUI)
		raise
	log("Done.", flush=True, level=LOW, GUI=GUI)

	log("Opening response spreadsheet...", end=' ', flush=True, level=LOW, GUI=GUI)
	sheet = gc.open_by_key(CONFIG["RESPONSES_SHEET"])
	responses_worksheet = sheet.worksheet_by_title("Responses")
	overview_worksheet = sheet.worksheet_by_title("Overview")
	log("Opened \"{}\"".format(sheet.title, GUI=GUI), flush=True, level=LOW, GUI=GUI)

	template = None

	now = str(datetime.datetime.now()).split('.')[0][:-3]
	log("Current time: {}".format(now), level=HIGH, GUI=GUI)
	#Actually read responses
	log("Reading all responses...", end=' ', flush=True, level=HIGH, GUI=GUI)
	cell_matrix = responses_worksheet.get_all_values(returnas='matrix')
	cell_matrix = [row for row in cell_matrix if row[0] != '']
	cell_matrix = cell_matrix[1:]
	log("Read all {} responses.".format(len(cell_matrix)), end=' ', flush=True, level=HIGH, GUI=GUI)
	diff = len(cell_matrix)-CONFIG["LAST_CHECKED_ENTRIES"]
	if(diff==0):
		log("No new responses.", level=HIGH, GUI=GUI)
	elif(diff==1):
		log("1 new response.", level=HIGH, GUI=GUI)
	else:
		log("{} new responses.".format(diff), level=HIGH, GUI=GUI)
	if(force):
		log("Force Updating all {} responses anyway.".format(len(cell_matrix)), level=HIGH, GUI=GUI)
	log("Processing entries into objects...", end=' ', flush=True, level=LOW, GUI=GUI)
	#Process said responses into objects
	people = {}
	updated_people = []
	for i,row in enumerate(cell_matrix):
		email = row[EMAIL]
		if(email not in people.keys()):
			try:
				name = names[email]
			except KeyError:
				name = email
			people[email] = Person(email, name)
			people[email].addLastCheck(now)
		people[email].addHours(row)
		if((force or i>=CONFIG["LAST_CHECKED_ENTRIES"]) and email not in updated_people):
			updated_people.append(email)
	log("Done.", flush=True, level=LOW, GUI=GUI)

	#We first sort by total number of hours, then by hours remaining.
	#This way people who are tied for hours remaining will be sorted
	#by their total hours, and people with twenty in hours but no out
	#hours will be below someone with ten and ten.
	log("Sorting People...", end=' ', flush=True, level=LOW, GUI=GUI)
	people = people.values()
	people = sorted(people, key=lambda person: person.getTotal())
	people = sorted(people[::-1], key=lambda person: person.getRemaining())
	log("Sorted.", flush=True, level=LOW, GUI=GUI)

	log("Finding individual sheet ids...", end=' ', flush=True, level=LOW, GUI=GUI)
	individual_sheets = {}
	individual_sheet_list = gc.list_ssheets(parent_id=CONFIG["INDIVIDUAL_SHEETS_DIR"])
	for sheet in individual_sheet_list:
		name = '\''.join(sheet['name'].split('\'')[:-1])
#		name = sheet['name'][:-(len(INDIVIDUAL_SHEETS_NAME.format(name='', year=CONFIG["YEAR"]))-4)]
		individual_sheets[name] = sheet['id']
	del individual_sheet_list
	log("Done.", flush=True, level=LOW, GUI=GUI)

	timeRemaining = None
	if(GUI is not None):
		GUI.setProgressRange(len(people))
		GUI.clearProgressBar("0/"+str(len(people)))
		changedPeople = len(updated_people)
		unchangedPeople = len(people)-len(updated_people)
		timeRemaining = UNCHANGED_DELAY * unchangedPeople + CHANGED_DELAY * changedPeople
		print("Changed:",changedPeople, "Unchanged:", unchangedPeople, "Time Remaining:", timeRemaining)

	#Update everyone's personal sheets
	for i,person in enumerate(people):
		if(GUI is not None):
			GUI.setProgress(i, timeRemaining, (CHANGED_DELAY if person.email in updated_people else UNCHANGED_DELAY))
#		person.addLastCheck(str(datetime.datetime.now()).split('.')[0])
		personal_sheet = None
		if(person.name in individual_sheets):
			personal_sheet = gc.open_by_key(individual_sheets[person.name])
		else:
			log("Creating new spreadsheet for {}...".format(person.name), end=' ', flush=True, level=HIGH, GUI=GUI)
			if(template is None):
				template = gc.open_by_key(CONFIG["PERSONAL_TEMPLATE"])
#			personal_sheet = gc.create(INDIVIDUAL_SHEETS_NAME.format(name=person.name, year=CONFIG["YEAR"]))
			personal_sheet = gc.create(INDIVIDUAL_SHEETS_NAME.format(name=person.name, year=CONFIG["YEAR"]), parent_id=CONFIG["INDIVIDUAL_SHEETS_DIR"])
			log("Setting up worksheets...", end=' ', flush=True, level=HIGH, GUI=GUI)
			for worksheet in ["Overview", "In Hours", "Out Hours"]:
				personal_sheet.add_worksheet(worksheet,src_worksheet=template.worksheet_by_title(worksheet))
			personal_sheet.del_worksheet(personal_sheet.sheet1)
			log("Sharing...", end=' ', flush=True, level=HIGH, GUI=GUI)
			#personal_sheet.share(person.email, role='reader')
#			file = gc.files().get(fileId=personal_sheet.id,fields='parents').execute()
#			previous_parents = ",".join(file.get('parents'))
#			file = gc.files().update(fileId=personal_sheet.id, addParents=CONFIG["INDIVIDUAL_SHEETS_DIR"], removeParents=previous_parents, fields="id, parents").execute()
#			print(previous_parents)
			individual_sheets[person.name] = personal_sheet.id
			log("Done!", flush=True, level=HIGH, GUI=GUI)
			time.sleep(4) #Don't hit the rate quota
		person.addSheet(personal_sheet.id)

		if(person.email in updated_people):
			log("Updating {}'s sheet...".format(person.name), end=' ', flush=True, level=MED, GUI=GUI)
			personal_sheet.worksheet_by_title("Overview").update_cells('A3:F3', [person.getPersonalOverview()])
			updateWorksheet(personal_sheet.worksheet_by_title("In Hours"), person.in_hours.getMatrix())
			updateWorksheet(personal_sheet.worksheet_by_title("Out Hours"), person.out_hours.getMatrix())
			log("Done!", flush=True, level=MED)
			time.sleep(CHANGED_DELAY) #Don't hit the rate quota
			timeRemaining -= CHANGED_DELAY
		else:
			log("Updating {}'s last checked time...".format(person.name), end=' ', flush=True, level=LOW, GUI=GUI)
			personal_sheet.worksheet_by_title("Overview").update_cell('F3', person.last_check)
			log("Done!", flush=True, level=LOW)
			time.sleep(UNCHANGED_DELAY) #Don't hit the rate quota
			if(GUI is not None):
				timeRemaining -= UNCHANGED_DELAY

	if(GUI is not None):
		GUI.setProgress(len(people), 0)

	if(len(updated_people) > 0):
		if(GUI is not None):
			GUI.pulseProgress()
		log("Constructing Overview Output Matrix...", end=' ', flush=True, level=HIGH, GUI=GUI)
		output_matrix = [person.getOverview() for person in people]
		log("Writing Overview...", end=' ', flush=True, level=HIGH, GUI=GUI)
		overview_worksheet.resize(rows=len(output_matrix)+2, cols=10)
		overview_worksheet.update_cells('A3:G{}'.format(len(output_matrix)+2), output_matrix)
		log("Done.", flush=True, level=HIGH, GUI=GUI)
		time.sleep(1) #Don't hit the rate quota

		CONFIG["LAST_CHECKED_ENTRIES"] = len(cell_matrix)
		writeConfig(CONFIG, DIR)
	else:
		log("Skipping updating overview, as nothing has changed", level=HIGH, GUI=GUI)

def openSheet():
	webbrowser.open("https://docs.google.com/spreadsheets/d/{}/".format(CONFIG["RESPONSES_SHEET"]))

if(__name__ == "__main__"):
	try:
		readConfig(sys.argv[1])
	except IndexError:
		readConfig()
	try:
		while(True):
			print("Starting")
			updateHours()
			break
	except KeyboardInterrupt:
		log("\nStopping.")
	except socket.timeout:
		log("\nTimed out. Restarting.")
	except:
		log("An error has occured at {}".format(str(datetime.datetime.now()).split('.')[0]))
		print("\n\nAn error has occured at {}".format(str(datetime.datetime.now()).split('.')[0]), file=sys.stderr)
		raise
	log("Done!")

