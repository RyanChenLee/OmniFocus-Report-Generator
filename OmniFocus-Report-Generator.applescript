(*
◸ RYANCHENLEE
OmniFocus - Weekly Project Report Generator
Authored by Ryan Lee
VERSION 1.0
December 28, 2020

// DESCRIPTION:
This script is a periodic OF report generator that prompts the user a time range and generates a formatted Evernote note of completed tasks sorted by project.   

// SOURCES USED:
-- https://joebuhlig.com/a-record-of-completed-tasks/
-- http://veritrope.com/code/omnifocus-weekly-project-report-generator
-- http://www.tuaw.com/2013/02/18/applescripting-omnifocus-send-completed-task-report-to-evernot/


// REQUIREMENTS
More details on the script information page.

//CHANGELOG
1.0 Initial Release

*)

(*
======================================
// MAIN PROGRAM
======================================
*)

-- Prepare a name for the new Evernote note
set theNoteName to "OF ✅ Report"
set theNotebookName to "📥 Inbox"

-- Prompt the user to choose a scope for the report
activate
set theReportScope to choose from list {"Today", "Yesterday", "This Week", "Last Week", "This Month", "Last Month", "Custom Range"} default items {"Yesterday"} with prompt "Generate a report for:" with title theNoteName
set theReportTitle to display dialog "Title of Report (optional)" as text default answer "" with title theNoteName

if theReportScope = false then return
set theReportScope to item 1 of theReportScope

(*
======================================
// OmniFocus: Grabbing Time
======================================
*)
-- Calculate the task start and end dates, based on the specified scope
set theStartDate to current date
set hours of theStartDate to 0
set minutes of theStartDate to 0
set seconds of theStartDate to 0
set theEndDate to theStartDate + (23 * hours) + (59 * minutes) + 59

if theReportScope = "Today" then
	set theDateRange to date string of theStartDate
else if theReportScope = "Yesterday" then
	set theStartDate to theStartDate - 1 * days
	set theEndDate to theEndDate - 1 * days
	set theDateRange to date string of theStartDate
else if theReportScope = "This Week" then
	repeat until (weekday of theStartDate) = Sunday
		set theStartDate to theStartDate - 1 * days
	end repeat
	repeat until (weekday of theEndDate) = Saturday
		set theEndDate to theEndDate + 1 * days
	end repeat
	set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
else if theReportScope = "Last Week" then
	set theStartDate to theStartDate - 7 * days
	set theEndDate to theEndDate - 7 * days
	repeat until (weekday of theStartDate) = Sunday
		set theStartDate to theStartDate - 1 * days
	end repeat
	repeat until (weekday of theEndDate) = Saturday
		set theEndDate to theEndDate + 1 * days
	end repeat
	set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
else if theReportScope = "This Month" then
	repeat until (day of theStartDate) = 1
		set theStartDate to theStartDate - 1 * days
	end repeat
	repeat until (month of theEndDate) is not equal to (month of theStartDate)
		set theEndDate to theEndDate + 1 * days
	end repeat
	set theEndDate to theEndDate - 1 * days
	
	set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
else if theReportScope = "Last Month" then
	if (month of theStartDate) = January then
		set (year of theStartDate) to (year of theStartDate) - 1
		set (month of theStartDate) to December
	else
		set (month of theStartDate) to (month of theStartDate) - 1
	end if
	set month of theEndDate to month of theStartDate
	set year of theEndDate to year of theStartDate
	repeat until (day of theStartDate) = 1
		set theStartDate to theStartDate - 1 * days
	end repeat
	repeat until (month of theEndDate) is not equal to (month of theStartDate)
		set theEndDate to theEndDate + 1 * days
	end repeat
	set theEndDate to theEndDate - 1 * days
	set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
	-- Custom Date Range
else if theReportScope = "Custom Range" then
	
	set date1 to display dialog "Enter start date in this format MM/DD/YY" as text default answer "" with title theNoteName
	set date2 to display dialog "Enter end date in this format MM/DD/YY" as text default answer "" with title theNoteName
	set theStartDate to date (text returned of date1)
	set theEndDate to (date (text returned of date2)) + (23 * hours) + (59 * minutes) + 59
	
	set theDialogStartText to "theStartDate is " & theStartDate & "."
	display dialog theDialogStartText
	
	-- Debugging: Output dates for entries
	-- set theDialogEndText to "theEndDate is " & theEndDate & "."
	-- display dialog theDialogEndText
	
	
	-- set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
	
	
end if
(*
======================================
// Omnifocus: Parsing tasks
======================================
*)
-- Begin preparing the task list as HTML
set theReportDate to do shell script "date +%Y-%m-%d"
set theProgressDetail to "<html><body><h1>Completed Tasks Report</h1></br><b>" & theDateRange & "</b><br>" & "Generated on: " & theReportDate & "<br><hr><br>"
set theInboxProgressDetail to "<br>"

-- Retrieve a list of projects modified within the specified scope
set modifiedTasksDetected to false
tell application "OmniFocus"
	tell front document
		set theModifiedProjects to every flattened project where its modification date is greater than theStartDate and modification date is less than theEndDate
		-- Loop through any detected projects
		repeat with a from 1 to length of theModifiedProjects
			set theCurrentProject to item a of theModifiedProjects
			-- Retrieve any project tasks modified within the specified scope
			set theCompletedTasks to (every flattened task of theCurrentProject where its completed = true and completion date is greater than theStartDate and completion date is less than theEndDate)
			-- Loop through any detected tasks
			if theCompletedTasks is not equal to {} then
				set modifiedTasksDetected to true
				-- Append the project name to the task list
				
				--ITERATION if want to print project completion
				--To-Do fix "missing value" issue when no due date
				-- set theCompletedProjectDate to completion date of theCurrentProject
				-- if theCompletedProjectDate is equal to "missing value" then set theCompletedProjectDate to "[Ongoing]"
				
				-- set theProgressDetail to theProgressDetail & "<h2>" & name of theCurrentProject & "</h2>" & "<h4>[" & theCompletedProjectDate & "]</h4>" & return & "<br><ul>"
				
				set theProgressDetail to theProgressDetail & "<h2>" & name of theCurrentProject & "</h2>" & return & "<br><ul>"
				repeat with b from 1 to length of theCompletedTasks
					set theCurrentTask to item b of theCompletedTasks
					set theCompletedDate to completion date of theCurrentTask
					set theCompletedTime to time string of theCompletedDate
					-- Append the tasks's name to the task list
					set theProgressDetail to theProgressDetail & "<li>" & "<b>" & name of theCurrentTask & "</b>" & "<br>" & " ----- " & theCompletedDate & "</li>" & return
				end repeat
				set theProgressDetail to theProgressDetail & "</ul>" & "<hr>" & return
			end if
		end repeat
		-- Include the OmniFocus inbox
		set theInboxCompletedTasks to (every inbox task where its completed = true and completion date is greater than theStartDate and completion date is less than theEndDate)
		-- Loop through any detected tasks
		if theInboxCompletedTasks is not equal to {} then
			set modifiedTasksDetected to true
			-- Append the project name to the task list
			set theInboxProgressDetail to theInboxProgressDetail & "<h2>" & "Inbox" & "</h2>" & return & "<br><ul>"
			repeat with d from 1 to length of theInboxCompletedTasks
				-- Append the tasks's name to the task list
				set theInboxCurrentTask to item d of theInboxCompletedTasks
				set theInboxProgressDetail to theInboxProgressDetail & "<li>" & name of theInboxCurrentTask & "</li>" & return
			end repeat
			set theInboxProgressDetail to theInboxProgressDetail & "</ul>" & return
		end if
		
	end tell
end tell
set theProgressDetail to theProgressDetail & theInboxProgressDetail & "</body></html>"

-- Notify the user if no projects or tasks were found
if modifiedTasksDetected = false then
	display alert "OmniFocus Completed Task Report" message "No modified tasks were found for " & theReportScope & "."
	return
end if

(*
======================================
// Evernote: Create report in note
======================================
*)

-- Create the note in Evernote.
tell application "Evernote"
	activate
	set theReportDate to do shell script "date +%Y-%m-%d"
	set theNote to create note notebook theNotebookName title "[" & theNoteName & "] " & theReportScope & ": " & text returned of theReportTitle & " [" & theReportDate & "]" with html theProgressDetail
	open note window with theNote
end tell