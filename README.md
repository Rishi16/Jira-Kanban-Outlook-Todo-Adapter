# Jira-Outlook-Todo-Adapter
Purpose - Used as an adapter to migrate outlook tasks to Jira.

Not all the issues you work on are logged in JIRA. Creating a custom issue in Jira is time consuming.
Mostly all of the issues you work on are present in the email chain. And you usually mark those emails as todo which you want to add in the tasklist.
This adapter fetches the subject of these todo email as summary and email body as descirption to create an issue on JIRA board.
This adapter provides you an UI to enter the Jira Assigne ID, Jira Login ID, Jira Password, Board name, Board ID and Jira link.
This adapter runs in background creating a system try icon.
You can either run it once or run it scheduled after every 3 minutes. So once you schedule this adapter you never have to worry about running it again until you restart your windows system.
This Adapter will keep your tasklist synced with the Jira board.


Features Implemented:
1. Populate Jira with outlook tasks
2. Check if the task is already present on Kanban board. If not, add.
3. Transit tasks from NS to DONE if marked completed in outlook.
4. Archive tasks from DONE to ARCHIVE which are older than a week.
5. Sync your Outlook tasklist to jira forever.
