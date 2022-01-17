# Attendance Summarizer - Robotic Process Automation (UiPath)


## Problem Statement/Use case Identification: 
Attendance tracking for online conferences/classes is a cumbersome process. There is a need for automation to achieve the following objectives:

1. To clean machine-generated attendance data.
2. Eliminate the redundant data entry process.
3. To present data in a simplified and well-formatted manner.

For instance, if a college instructor wishes to update the attendance of students who have attended his/her online class but has to do the repetitive task of checking whether the student’s name/email is present in the attendance CSV file, he/she may easily do so by using the attendance - summarizer bot.


## Introduction:
From online classes to international conferences, the attendance of virtual meetings can be better monitored using this automation process.
The data filtered can be used in various processes such as updating attendance on online portals of institutions, generating certificates of participation, and so on. 
We have used two excel sheets for this project, the input sheet which gives a list of all the attended participants (a subset of all the students), duration of their attendance and other details; and the output sheet, which compiles the attendance for the universal set of students. 

Excel application scope activity holds all the activity that takes place in the attendance summarizer project. Using the Read range activity, we read all the data (email ids) from the input file. For Each Row in Data Table is used to iterate through each email id, parse it and clean the data to the required format, that is 1MS*******. This is now assigned to a new variable and using the LookUp Range activity in the output sheet USN list, we check if the USN exists, which indicates the student is present, else absent.

## Designing the RPA Bot: 

Step 1: Create a new process “ Attendance Summariser”.
Step 2: Create a new sequence.

Step 3: Drag and drop Excel application scope activity.
 	Set up an Excel application scope with the Excel file with student details.

Step 4: Drag and drop the Read range activity and specify the sheet to be read.
	Create variable `outputDataTable` of type `Datatable`

Step 5:  Drag and drop read CSV to read from the attendance CSV file and to output the data into 
              `outputDataTable` variable.
Step 6:  Drag and drop For Each Row in Data Table activity.

Step 7:  Create variable `x` of type String, `usnRaw` of type String, `first1ms` of type Int32.

Step 8:  “For each row in outputDataTable”
Assign:
```
x=x.Split({"@"},stringsplitoptions.None).ToList(0)
usnRaw = x.Split({"@"},stringsplitoptions.None).ToList(0)
first1ms = usnRaw.IndexOf("1MS")
```
Step 9:  Create variable `usn` of type String, `studentSheet` of type DataTable

Step 10: Drag and drop If- else activity 
Set condition : `first1ms>-1 And usnRaw.Length>=10`
Then:
	Drag and drop Assign activity 
		Assign:
    ```
    usn = usnRaw.Substring(first1ms,10)
    ```
	Drag and drop LookUp Range activity 
		Set sheet name and value to be searched as “usn”
	Drag and drop Lookup Data Table activity 
		Set Column Index as the per the column which contains the USN 
		(in our case , column index =1)
		Set `DataTable= studentSheet`
		Set `LookupValue = usn`
	Drag and drop another If-else Activity
		Condition : `Not usnLocation.Equals("-1")`
		Then:
			Drag and drop Write Cell activity
				Set sheetname accordingly (ex: “sheet1”)
				Set value to be entered as “Yes” ( presentees as Yes)
Set Range as `"F"+usnLocation.Substring(1)`

## Applications /Usage of RPA Bot. 
- In educational institutions like schools and colleges, teachers can automate the process of marking and keeping track of the attendance of students
- It also improves the accuracy of the process and provides ease in workflow management, reducing hassle in the teaching process
- Can eliminate proxy attendances, caused while taking traditional roll call in online classes.
- Companies organizing conferences can assess the participation of their employees
- Used in marketing to evaluate the customers' preferences and interests, based on their engagement with the product
- To keep track of customer retention as a whole
- Online events and webinars can be monitored to check for invalid users joining.

## Conclusion:
The manual process of attendance tracking and evaluation which can seem quite cumbersome with the online classes can now be solved with one click of the UIPath Automation bot, “attendance-summarizer”. Furthermore, the bot can be customized to analyze student engagement and analyze student performance using attendance data.
 
