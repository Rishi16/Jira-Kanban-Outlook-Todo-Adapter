'''
Nov-2017
Author - Rishikesh Shendkar (rishikesh.shendkar@gmail.com)

Purpose - Used as an adapter to migrate outlook tasks to Jira.

Features Implemented:
1. Populate Jira with outlook tasks
2. Check if the task is already present on Kanban board. If not, add.
3. Transit tasks from NS to DONE if marked completed in outlook.
4. Archive tasks from DONE to ARCHIVE which are older than a week.

'''

# Official Python module from Jira
from jira import JIRA
from jira.exceptions import JIRAError
# In-built Python logging module
import logging
# In-built Python module for sys.exit()
import sys
# Official Python module from windows for accessing certain windows applications
import win32com.client
# In-built Python module for suppressig the insecure request warning.
import urllib3
import re

# This function is the main function running other functions
def syncTasksToJira(jiraID, jiraUsername, jiraPassword, boardName, boardID, jiraLink):
	# Configuration Start xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
	LOGFORMAT= '[%(asctime)s - %(levelname)s: %(funcName)20s()] %(message)s'
	logging.basicConfig(filename="jira-adapter.log",level = logging.INFO, format = LOGFORMAT)
	logger = logging.getLogger('JiraOutAdapter')
	urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

	options = {
		'server' : jiraLink,
		'verify' : False,
	}
	
	# Establishing the connecetio to JIRA
	try:
		jira = JIRA(options=options, basic_auth=(jiraUsername, jiraPassword))
	except JIRAError as jex:
		errorText = jex.text.split(';')
		logger.exception('\t\t[JIRA EXCEPTION] - Connection Failure {0} - {1}\n'.format(jex.status_code, errorText[0]))
		print('\t\t[JIRA EXCEPTION] - Connection Failure {0} - {1}\n'.format(jex.status_code, jex.text))
		sys.exit(-1)
	except Exception as ex:
		logger.exception('\t\t[EXCEPTION] - Connection Failure - {0}'.format(ex))
		sys.exit(-1)
	# Initialising Board
	BOARD = {}
	BOARD['ID'] = boardID 
	BOARD['Name'] = boardName
	# Fetching the outlook tasks
	olFolderTodo = 28
	outlook = win32com.client.Dispatch("Outlook.Application")
	ns = outlook.GetNamespace("MAPI")
	todo_folder = ns.GetDefaultFolder(olFolderTodo)
	todo_items = todo_folder.Items
	tasks = todo_items.Restrict("[Complete] = FALSE")
	
	logger.info('Found {0} tasks for {1} board.'.format(len(todo_items), "BOARD"))
	# Initialising config values
	defaulttaskvalues = {}
	defaulttaskvalues['assigneeID'] = jiraID
	defaulttaskvalues['labels'] = ['OutlookTasks']

	for taskSN, task in enumerate(tasks):
		task.Subject = cleanse(task.Subject)
		print(u'\tToDo Task {0}: {1} '.format(taskSN, task.Subject))
		logger.info(u'ToDo Task {0}: {1} '.format(taskSN, task.Subject))
		existingIssue = get_existing_workitem(jira, BOARD['ID'], task, defaulttaskvalues, customJQL=None)
		if not existingIssue:
			# This creates a new Incident work item on the board
			newIssue = create_workitem_tasks(jira, BOARD['ID'], task, defaulttaskvalues)
			#task.Subject = u'{1} -{0} '.format(newIssue, task.Subject)
			#task.Save()
	
	archive_tasks_from_done_stage(jira, BOARD['ID'], defaulttaskvalues)
	transit_tasks_to_done_stage(jira, BOARD['ID'], defaulttaskvalues, todo_items)
	print('Synced')
	logger.info('Synced')	
	# We are done. kthnxbye


# This function archives DONE tasks older than a week
def archive_tasks_from_done_stage(jira, project, defaulttaskvalues):
	logger = logging.getLogger('JiraOutAdapter')
	print("\nArchiving Tasks")
	# Custom JQL to get all the issues in Done stage older than a week
	customJQL = "project={0} and assignee={1} and labels={2} and status in (Done) and createdDate <= -1w".format( str(project), defaulttaskvalues['assigneeID'], "".join(defaulttaskvalues['labels']))
	try:
		issues = jira.search_issues(customJQL, startAt=0, maxResults=1000)
		if not issues:
			logger.info('\TFound {0} existing issues. Checking again...'.format(len(issues)))
			issues = jira.search_issues(customJQL, startAt=0, maxResults=1000)
		
		if issues:
			logger.info('\tFound {0} issues in Done stage. Checking for archive transition.'.format(len(issues)))
			for i, issue in enumerate(issues):
				print('\tArchiving [{0}] {1}'.format(i, issue))
				jira.transition_issue(issue, 'Done to Archive')
				logger.info('Issue {0} has been archived '
								'in Kanban board.'.format(issue))
	except JIRAError as jex:
		logger.exception('\t\t[JIRA EXCEPTION] - Archive Issues {0} - {1}\n'.format(jex.status_code, jex.text))
		print('\t\t[JIRA EXCEPTION] - Archive Issues {0} - {1}\n'.format(jex.status_code, jex.text))
	except Exception as ex:
		logger.exception('\t\t[EXCEPTION] - Archive Search Issue {0}'.format(ex))
		print(ex)
	# Done

def cleanse(line):
	line=line.replace('FW: ', '')
	line=line.replace('RE: ', '')
	line=re.sub(r'([^\s\w]|_)+', '', line)
	line=line.replace('   ', ' ')
	line=line.replace('  ', ' ')
	return line	

# This method helps to find an existing tasks work item in Jira based on the task subject	
def get_existing_workitem(jira, project, task, defaulttaskvalues, customJQL):
	logger = logging.getLogger('JiraOutAdapter')
	# If this parameter was not passed, then assume we need to check whole of the project.
	if not customJQL:
		customJQL = "project={0} and assignee={1} and labels={2} and summary ~ '{3}' and " \
						"status not in (Closed, Archive)".format(str(project),defaulttaskvalues['assigneeID'], "".join(defaulttaskvalues['labels']), task.Subject)
	# to reduce the number of items returned, we can further narrow down the search using more filter parameters.
	# There is a restriction of 1000 items on search function.
	else:
		customJQL = customJQL + ' and project=' + str(project)

	# Holds the issue found through this method
	issueExisting = 'None'
	try:
		issues = jira.search_issues(customJQL, startAt=0, maxResults=1)
		if not issues:
			logger.info('\tFound {0} existing issues. Checking again...'.format(len(issues)))
			issues = jira.search_issues(customJQL, startAt=0, maxResults=1)

		if issues:
			if len(issues) >= 1:
				print('\tFound {0} existing issue for {1}'.format(issues[0], task.Subject))
				logger.info('\tFound {0} existing issue for {1}'.format(issues[0], task.Subject))
				issueExisting = issues[0]
		else:
			issueExisting = None

	except JIRAError as jex:
		logger.exception('\t\t[JIRA EXCEPTION] - Search Issue {0} - {1}\n'.format(jex.status_code, jex.text))
		print('\t\t[JIRA EXCEPTION] - Search Issue {0} - {1}\n'.format(jex.status_code, jex.text))
	except Exception as ex:
		logger.exception('\t\t[EXCEPTION] - Search Issue {0}'.format(ex))
		print(ex)

	return issueExisting

# This function creates the tasks send to it.	
def create_workitem_tasks(jira, project, task, defaulttaskvalues):
	logger = logging.getLogger('JiraOutAdapter')
	# Holds the new issue ID from Jira when created successfully
	new_issue = None
	# Holds the parameters for the JSON for creating a new issue via API
	issue_dict = {}
	defaulttaskvalues['Priority'] = 'Medium'

	# Gets Project ID
	if project:
		issue_dict['project'] = {'id' : project}
	else:
		logger.warning('No project ID was passed.')
	# Gets Summary. This will show on the cards on the board
	issue_dict['summary'] = '{0}'.format(task.Subject.replace('\n', ''))

	# Gets the Notes
	if task.Body:
		issue_dict['description'] = task.Body
	else:
		logger.warning('No task note was passed to {0}'.format(task.Subject))

	# Gets Priority info.
	if defaulttaskvalues['Priority']:
		issue_dict['priority'] = {'name' : defaulttaskvalues['Priority']}
	else:
		logger.warning('No priority was passed. {0}'.format(task.Subject))

	# Sets the issue type
	issue_dict['issuetype'] = {'name': 'Story'}
	issue_dict['components'] = [{'name' : 'Maintenance Tasks'}]
	if defaulttaskvalues['labels']:
		issue_dict['labels'] = defaulttaskvalues['labels']
	else:
		logger.warning('No label was passed. {0}'.format(task.Subject))

	# Gets assignee
	if defaulttaskvalues['assigneeID']:
		issue_dict['assignee'] ={'name' : defaulttaskvalues['assigneeID']}
	else:
		logger.warning('No assignee ID was passed. {0}'.format(task.Subject))

	try:
		# Pass all the collected info to Jira's API to create the issue on Kanban board
		new_issue = jira.create_issue(fields=issue_dict)
		print('\tCreated issue - {0}\n'.format(new_issue))
		logger.info('\tCreated issue - {0} for task {1}'.format(new_issue, task.Subject))
	except JIRAError as jex:
		logger.info(issue_dict)
		logger.exception('\t\t[JIRA EXCEPTION] Create issue - {0} - {1}\n'.format(jex.status_code, jex.text))
		print('\t\t[JIRA EXCEPTION] Create issue - {0} - {1}\n'.format(jex.status_code, jex.text))
	except Exception as ex:
		logger.info(issue_dict)
		logger.exception('\t\t[EXCEPTION] Create issue - {0}'.format(ex))
		print(ex)

	return new_issue

# This function transits completd tasks in outlook to DONE stage
def transit_tasks_to_done_stage(jira, project, defaulttaskvalues, todo_items):
	logger = logging.getLogger('JiraOutAdapter')
	# Fecthing Completed tasks only
	#issue = jira.issue(project)
	#transitions = jira.transitions(issue)
	#[(t['id'], t['name']) for t in transitions]
	print('\nTransiting Completed Tasks')
	tasks = todo_items.Restrict("[Complete] = TRUE")
	for i,task in enumerate(tasks):
		task.Subject = cleanse(task.Subject)
		print("\t[{0}] {1}".format(i,task.Subject))
		customJQL = "project={0} and assignee={1} and labels={2} and summary ~ '{3}' and " \
						"status not in (Closed, Archive)".format(str(project),defaulttaskvalues['assigneeID'], "".join(defaulttaskvalues['labels']), task.Subject)
		try:
			issues = jira.search_issues(customJQL, startAt=0, maxResults=10)
			if not issues:
				logger.info('\tFound {0} existing issues. Checking again...'.format(len(issues)))
				issues = jira.search_issues(customJQL, startAt=0, maxResults=10)
	
			if issues:
				for i, issue in enumerate(issues):
					try:
						print("\t\t Transitioning from NS to WIP")
						jira.transition_issue(issue, transition='Move From NS to WIP')
						print("\t\t Transitioned from NS to WIP")
					except JIRAError as jex:
						logger.exception('\t\t[JIRA EXCEPTION] - {2} - NS to WIP - {0} - {1}\n'.format(jex.status_code, jex.text, issue))
						print('\t\t[JIRA EXCEPTION] - {2} - Transition to Done - {0} - {1}\n'.format(jex.status_code, jex.text, issue))
					except Exception as ex:
						logger.exception('\t\t[EXCEPTION] Transition to Done - {0}'.format(ex))
						print('\t\t[EXCEPTION] Transition to Done - {0}'.format(ex))
					try:
						print("\t\t Transitioning from Deffered to WIP")
						jira.transition_issue(issue, transition='Deferred to WIP')
						print("\t\t Transitioned from Deffered to WIP")
					except JIRAError as jex:
						print('\t\t[JIRA EXCEPTION] - {2} - Transition to Done - {0} - {1}\n'.format(jex.status_code, jex.text, issue))
						logger.exception('\t\t[JIRA EXCEPTION] - {2} - Deferred to WIP- {0} - {1}\n'.format(jex.status_code, jex.text, issue))
					except Exception as ex:
						logger.exception('\t\t[EXCEPTION] Transition to Done - {0}'.format(ex))
						print('\t\t[EXCEPTION] Transition to Done - {0}'.format(ex))
					try:
						print("\t\t Transitioning from WIP to READY")
						jira.transition_issue(issue, transition='WIP to Ready')
						print("\t\t Transitioned from WIP to READY")
					except JIRAError as jex:
						logger.exception('\t\t[JIRA EXCEPTION] - {2} - WIP to Ready - {0} - {1}\n'.format(jex.status_code, jex.text, issue))
						print('\t\t[JIRA EXCEPTION] - {2} - Transition to Done - {0} - {1}\n'.format(jex.status_code, jex.text, issue))
					except Exception as ex:
						logger.exception('\t\t[EXCEPTION] Transition to Done - {0}'.format(ex))
						print('\t\t[EXCEPTION] Transition to Done - {0}'.format(ex))
					try:
						print("\t\t Transitioning from Ready to Done")
						jira.transition_issue(issue, transition='Ready to Done')
						print("\t\t Transitioned from Ready to Done")
					except JIRAError as jex:
						logger.exception('\t\t[JIRA EXCEPTION] - {2} - Ready to Done - {0} - {1}\n'.format(jex.status_code, jex.text, issue))
						print('\t\t[JIRA EXCEPTION] - {2} - Transition to Done - {0} - {1}\n'.format(jex.status_code, jex.text, issue))
					except Exception as ex:
						logger.exception('\t\t[EXCEPTION] Transition to Done - {0}'.format(ex))
						print('\t\t[EXCEPTION] Transition to Done - {0}'.format(ex))
					logger.info('Issue {0} has been archived '
									'in Kanban board.'.format(issue))	
		except JIRAError as jex:
			logger.exception('\t\t[JIRA EXCEPTION] Transition to Done - {0} - {1}\n'.format(jex.status_code, jex.text, issue))
			print('\t\t[JIRA EXCEPTION] Transition to Done - {0} - {1}\n'.format(jex.status_code, jex.text, issue))
		except Exception as ex:
			logger.exception('\t\t[EXCEPTION] Transition to Done - {0}'.format(ex))
			print('\t\t[EXCEPTION] Transition to Done - {0}'.format(ex))	
	# Done
