from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from email.mime.text import MIMEText
from datetime import datetime as dt
import base64
import os
import pickle
import json

# Define base directory
BASE_DIR = '/mnt/c/Users/garri/Documents/Projects/GoogleReminders/'
# Define path to JSONs saved for visualization
EXAMPLE_DIR = BASE_DIR+'exampleDocs/'
# Define info dir location
INFO_DIR = BASE_DIR+'info/'
# Path to your OAuth 2.0 credentials file
CREDENTIALS_FILE = BASE_DIR+'auth/RemindersCreds.json'

def append_to_sheet(service, sheetID, range, values):
	"""Appends values as rows to a sheet"""
	try:
		request = service.spreadsheets().values().append(
			spreadsheetId=sheetID,
			range=range,
			valueInputOption="RAW",
			insertDataOption="INSERT_ROWS",
			body={
				'values': values
			}
		)
		
		response = request.execute()
		print(f"Rows appended: {response.get('updates').get('updatedRows')}")
	except Exception as e:
		print(f"An error occurred: {e}")

def authenticate(scopes):
		"""Authenticate and create a service object for Google APIs."""
		creds = None
		if os.path.exists(BASE_DIR+'auth/'+'token.pickle'):
				with open(BASE_DIR+'auth/'+'token.pickle', 'rb') as token:
						creds = pickle.load(token)

		if not creds or not creds.valid:
				if creds and creds.expired and creds.refresh_token:
						creds.refresh(Request())
				else:
						flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, scopes)
						creds = flow.run_local_server(port=0)
				with open(BASE_DIR+'auth/'+'token.pickle', 'wb') as token:
						pickle.dump(creds, token)

		return creds

def display_document(service, title, filename):
	"""Save contents of Google Doc"""

	docIDs = get_documentIDs()
	doc = get_document(service, docIDs[title])

	with open(EXAMPLE_DIR + filename, 'w') as file:
		json.dump(doc, file, indent=2)

def get_checklists(service, docIDs, retry):
	"""Get checklists from all docs with ids in info/docIDs.json"""

	# Define checklist, list of tables not titled "Checklist: ", and list of invalid rows from last run
	checklists = []
	ignoredTables = []
	invalidRows = []
	lastTime = get_last_runtime('update_checklist')

	# Initialize list 
	rowsToRetry = []
	if retry:
		with open(INFO_DIR + 'failedRows.json', 'r') as file:
			rowsToRetry = json.load(file)

	# Append checklists and ignored tables through all checklist docs
	for title in docIDs:
		if not retry:
			print(f"\n== Processing {title} ==")
		else:
			isDocPresent = any(r["docID"] == docIDs[title] for r in rowsToRetry)
			if not isDocPresent:
				continue
			print(f"\n== Retrying {title} ==")
			
		doc = get_document(service, docIDs[title])

		# Get all checklists from doc
		docChecklists, docIgnoredTables, docInvalidRows = get_doc_checklists(doc, lastTime, docIDs[title], title, rowsToRetry)
		if docChecklists:
			checklists.extend(docChecklists)
		if docIgnoredTables:
			ignoredTables.extend(docIgnoredTables)
		if docInvalidRows:
			invalidRows.extend(docInvalidRows)

	with open(INFO_DIR + 'runtimes.json', 'r+') as file:
		runtimes = json.load(file)
		runtimes = {}
		runtimes['update_checklist'] = dt.now().strftime('%m-%d-%Y-%H-%M')
		file.seek(0)
		json.dump(runtimes, file, indent=2)
		file.truncate()

	with open(INFO_DIR + 'failedRows.json', 'w') as file:
		json.dump(invalidRows, file, indent=2)

	if invalidRows:
		print('\n\n')
	for row in invalidRows:
		print(f"Please fix table: {row["tableTitle"]} in: {row["link"]}")
	if invalidRows:
		print("\nThen input '2' to retry failed rows.")

	return checklists

def get_doc_checklists(doc, lastTime, docID, docTitle, rowsToRetry, tabID = '', tabTitle = ''):
	"""Get checklists from one doc or parent tab"""

	# Define checklists and ignored tables for doc
	docChecklists = []
	docIgnoredTables = []
	docInvalidRows = []

	# Determine the correct key for tabs based on the structure
	if 'tabs' in doc:
		tabKey = 'tabs'
	elif 'childTabs' in doc:
		tabKey = 'childTabs'
	else:
		print(" ---- WARNING ---- ")
		print(f"No tabs or childTabs found in doc: {docTitle}")
		return docChecklists, docIgnoredTables, docInvalidRows # return empty lists if no 'tabs' or 'childTabs'

	# Iterate through tabs (or childTabs) in doc or parent tab
	for tab in doc[tabKey]:
		if 'documentTab' not in tab or 'body' not in tab['documentTab'] or 'content' not in tab['documentTab']['body']:
			continue

		# Get tab info
		tabID = tab['tabProperties']['tabId']
		tabTitle = tab['tabProperties']['title']

		# Skip if retrying and tab not present in rowsToRetry
		isTabPresent = True
		if rowsToRetry:
			isTabPresent = any(r["tabID"] == tabID for r in rowsToRetry)

		# Get checklists from parent tab
		if isTabPresent:
			parentChecklists, parentIgnoredTables, parentInvalidRows = get_tab_checklists(tab, lastTime, docID, docTitle, tabID, tabTitle, rowsToRetry)
			if parentChecklists:
				docChecklists.extend(parentChecklists)
			if parentIgnoredTables:
				docIgnoredTables.extend(parentIgnoredTables)
			if parentInvalidRows:
				docInvalidRows.extend(parentInvalidRows)

		# If this tab has childTabs, recursively call get_doc_checklists for each child
		if 'childTabs' in tab:
			# Recursively get checklists for child tabs
			childChecklists, childIgnoredTables, childInvalidRows = get_doc_checklists(tab, lastTime, docID, docTitle, rowsToRetry, tabID, tabTitle)
			if childChecklists:
				docChecklists.extend(childChecklists)
			if childIgnoredTables:
				docIgnoredTables.extend(childIgnoredTables)
			if childInvalidRows:
				docInvalidRows.extend(childInvalidRows)

	return docChecklists, docIgnoredTables, docInvalidRows

def get_documentIDs():
	"""Get JSON file containing checklist document keypairs { title: ID }"""

	if os.path.exists(INFO_DIR+'checklistDocIDs.json'):
		with open(INFO_DIR+'checklistDocIDs.json', 'r') as file:
			docIDs = json.load(file)
	else:
		docIDs = {}

	return docIDs

def get_document(service, docID):
		"""Retrieve, print, and return the full content of a Google Docs document."""
		doc = service.documents().get(documentId=docID, includeTabsContent=True).execute()
		return doc

def get_last_runtime(name):
	"""Get last runtime of 'name' attr"""
	with open(INFO_DIR + 'runtimes.json', 'r') as file:
		times = json.load(file)
	
	if not times or not times[name]:
		print("Could not find value for key: '{name}' in runtimes.json")
	else:
		return times[name]

def get_tab_checklists(tab, lastTime, docID, docTitle, tabID, tabTitle, rowsToRetry):
	"""Get the checklists from one tab"""
	# Define dict of checklists from one tab
	tabChecklists = []
	tabIgnoredTables = []
	tabInvalidRows = []

	# Iterate through tab content looking for tables
	for tabElement in tab['documentTab']['body']['content']:
		# Skip tab element if no table
		if 'table' not in tabElement or 'tableRows' not in tabElement['table']:
			continue

		# Initialize title flag to break loop if table is not titled 'Checklist: ...'
		isChecklist = True

		# Initialize empty table title
		tableTitle = ''
		rowsAdded = 0
		# Iterate through rows of table
		for i, row in enumerate(tabElement['table']['tableRows']):
			# Skip second row  and row if no tableCells
			if i == 1 or 'tableCells' not in row:
				continue
			# Break out of loop if table not titled "Checklist: ..."
			if not isChecklist:
				break
			# If rowsToRetry exists (find tableTitle first)
			if i > 0 and rowsToRetry:
				# Break out of loop if tableTitle in tab and doc not present
				isTablePresent = any(r["docID"] == docID and r["tabID"] == tabID and r["tableTitle"] == tableTitle for r in rowsToRetry)
				if not isTablePresent:
					break
				# Skip current row if row in tableTitle, tab and doc not present
				isRowPresent = any(r["docID"] == docID and r["tabID"] == tabID and r["tableTitle"] == tableTitle and r["row"] == i for r in rowsToRetry)
				if not isRowPresent:
					continue

			# Initialize list to be appended to 'Checklist' sheet with link to doc
			checklistRow = [
				'https://docs.google.com/document/d/' + docID + '/edit?tab=' + tabID, 
				"'" + tableTitle + "' in '" + tabTitle + "' in '" + docTitle + "'"
			]
			# Initialize flag to guarantee invalid row is not appended
			isRowValid = True
			isContentPresent = True
			# Initialize flag to guarantee row with write date before last run date is not appended
			isNewRow = True

			# Iterate through cells of row
			for j, cell in enumerate(row['tableCells']):
				# Skip all but first cell in first row and any cells that do not contain 'elements'
				if 	(i == 0 and j != 0) or \
						'content' not in cell or \
						'paragraph' not in cell['content'][0] or \
						'elements' not in cell['content'][0]['paragraph']:
					continue

				# Initialize string to add cell content to
				outString = ''
				# Iterate through all elements in cell
				for element in cell['content'][0]['paragraph']['elements']:
					# Skip all elements with no textRun
					if 'textRun' not in element or 'content' not in element['textRun']:
						continue
					# Remove whitespace characters (spaces, '\n', '\t') and add a newline
					outString += element['textRun']['content'].strip() + '\n'
				# Remove last newline
				outString = outString.strip()

				# Get title and make sure 'Checklist...' is present in title
				if i == 0:
					fullTitle = outString.strip().split(':')
					tableTitle = ':'.join(fullTitle[1:]).strip()
					if fullTitle[0] != 'Checklist':
						tabIgnoredTables.append(f"Table: {outString} in tab: {tabTitle}, doc: {docTitle} ({checklistRow[0]})")
						isChecklist = False
						break
					elif not tableTitle:
						print("\n ---- UNTITLED TABLE WARNING ---- ")
						print(f"Untitled table: {outString} in tab: {tabTitle} in doc: {docTitle}")
				# Define checklist title, then keys, then fill with content
				else:
					# Add title to task column value if available
					if j == 0:
						if outString.strip():
							checklistRow.append(outString)
						else:
							checklistRow.append('')
					# Make sure valid date present then add to checklist if written since last run
					else:
						if j != 1:
							checklistRow.append(outString)
						try:
							checklistDateTime = dt.strptime(outString, '%m-%d-%Y-%H-%M')
						except ValueError:
							# Only warn if task present or datetime not empty
							if (checklistRow[2].strip() or outString.strip()):
								print("\n ---- INVALID DATETIME WARNING ---- " )
								print(f"Cell row: {i}, column: {j} in table: {tableTitle} in tab: {tabTitle} in doc: {docTitle} contains invalid datetime: '{outString}'.")
							else:
								isContentPresent = False
							isRowValid = False
							break
						# Break if written before last run and not retrying (already added)
						lastRunDateTime = dt.strptime(lastTime, '%m-%d-%Y-%H-%M')
						if not rowsToRetry and j == 1 and checklistDateTime < lastRunDateTime:
							isNewRow = False

			# Add 0 (incomplete status)
			checklistRow.append(0)
			# Append checklist row if current row write date is after last run date AND correct length and dates are valid
			if len(checklistRow) == 5 and isRowValid and isContentPresent and isNewRow:
				rowsAdded += 1
				tabChecklists.append(checklistRow)
			# Else if write date is after last run date add row to list for retry
			elif i != 0 and isNewRow and isContentPresent:
				tabInvalidRows.append({
					"docID": docID,
					"tabID": tabID,
					"tableTitle": tableTitle,
					"row": i,
					"link": checklistRow[0]
				})
				if len(checklistRow) != 5:
					print(f"\n ---- INVALID ROW LENGTH WARNING ---- ")
					print(f"Incorrect length row {i} in table: {tableTitle}")
					print(checklistRow)

		if rowsAdded > 0:
			print(f"{str(rowsAdded)} rows added from table: {tableTitle} in tab: {tabTitle}")

	return tabChecklists, tabIgnoredTables, tabInvalidRows

def update_checklist(docService, sheetService, retry):
	"""Update 'To Do' tab in 'Checklist' Google Sheet."""

	# Get all rows with valid date in second column from tables titled 'Checklist:...' in the docs with IDs in info/checklistDocIDs.json
	docIDs = get_documentIDs()
	newChecklists = get_checklists(docService, docIDs, retry)

	# Append rows to 'Checklist' spreadsheet
	with open(INFO_DIR+'destIDs.json', 'r') as file:
		destIDs = json.load(file)
	append_to_sheet(sheetService, destIDs['Checklist'], 'To Do', newChecklists)


# FOR REFERENCE ONLY
def create_document(service, title):
	"""Create a Google Doc with a custom title"""
	doc_body = {'title': title}
	doc = service.documents().create(body=doc_body).execute()

	# Add new docID
	# with open(INFO_DIR + 'docIDs.json', 'r+') as file:
	# 	docIDs = json.load(file)
	# 	docIDs[title] = doc['documentId']
	# 	file.seek(0)
	# 	json.dump(docIDs, file, indent=2)
	# 	file.truncate()

	# Save contents of new documents to example dir
	# with open(EXAMPLE_DIR + title + '.json', 'w') as file:
	# 	json.dump(doc, file, indent=2)

	print(f"Created doc '{title}' with ID: {doc['documentId']}")

	return doc['documentId']

def create_checklist_doc(service, title):
	"""Create a Google Doc using the checklist template"""

	docID = create_document(service, title)

	# Insert Checklist Table Request
	insertTable = [
		{
			'insertTable': {
				'rows': 3,
				'columns': 3,
				'location': { 'index': 1 },
			}
		}
	]

	'''
	# Insert Rows Request
	insertRows = [
		{
			'insertTableRow': {
				'tableCellLocation': {
					'tableStartLocation': { 'index': 2, 'tabId': 't.0' },
					'rowIndex': 0,
					'columnIndex': 0
				},
			"insertBelow": False
			}
		},
		{
			'insertTableRow': {
				'tableCellLocation': {
					'tableStartLocation': { 'index': 2, 'tabId': 't.0' },
					'rowIndex': 0,
					'columnIndex': 0
				},
			"insertBelow": False
			}
		}
	]
	'''

	# Alter Content of Table Request
	# First cell ('Checklist: '): 5 + 10 (text) + 2 (structural) = 23.
	# Second cell ('Task'): 23 + 4 (text) + 2 (structural) = 29.
	# Third cell ('Write Date'): 29 + 10 (text) + 2 (structural) = 41.
	insertContent = [
		{
			'mergeTableCells': {
				'tableRange': {
						'tableCellLocation': {
							'tableStartLocation': { 'index': 2, 'tabId': 't.0' },
							'rowIndex': 0,
							'columnIndex': 0
					},
					"rowSpan": 1,
					"columnSpan": 3
				}
			}
		},
		{
			'insertText': {
				'location': { 'index': 5 },
				'text': 'Checklist: '
			}
		},
		{
			'insertText': {
				'location': { 'index': 23 },
				'text': 'Task'
			}
		},
		{
			'insertText': {
				'location': { 'index': 29 },
				'text': 'Write Date'
			}
		},
		{
			'insertText': {
				'location': { 'index': 41 },
				'text': 'Due Date'
			}
		}
	]

	try:
		service.documents().batchUpdate(
			documentId=docID,
			body={'requests': insertTable}
		).execute()

		print('Inserted table.')
		try:
			service.documents().batchUpdate(
				documentId=docID,
				body={'requests': insertContent}
			).execute()

			print(f"Inserted content.")
		except Exception as e:
			print(f"Error updating rows: {e}\n\n")
	except Exception as e:
		print(f"Error inserting table: {e}\n\n")
