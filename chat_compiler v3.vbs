REM Nicholas Devlin
REM 08/13/2013
REM This code reads in message history then creates a new file for each month
REM Changes: variables instantiate automatically, allow for more than 2 people (not currently working)
' NEW LINE = vbcrlf

months = array("January","February","March","April","May","June","July","August","September","October","November","December")
'FUNCTIONS
function monthToNum(givenMonth)
select case givenMonth
case "January"
	monthToNum = "01"
case "February"
	monthToNum = "02"
case "March"
	monthToNum = "03"
case "April"
	monthToNum = "04"
case "May"
	monthToNum = "05"
case "June"
	monthToNum = "06"
case "July"
	monthToNum = "07"
case "August"
	monthToNum = "08"
case "September"
	monthToNum = "09"
case "October"
	monthToNum = "10"
case "November"	
	monthToNum = "11"
case "December"
	monthToNum = "12"
case else 
	monthToNum = "-1"
end select
end function


'READ/ PROCESS
totalFileName = InputBox("Enter the name of the text file. (do not include '.txt')"&VbCrLf&"File Name: ")
set objFSO = CreateObject("Scripting.FileSystemObject")
set objFile = objFSO.OpenTextFile(totalFileName & ".txt", 1) 'Open file to read


'create list of people in conversation
redim nameList(0) 'array of names
tempNameCount = 0
tempName = ""

textLine = objFile.ReadLine 'reads in first line which should always countain list of names seperated by commas
for i = 1 to len(textLine)
	myChar = Mid(textLine, i, 1)
	if(myChar <> ",") then
		tempName = tempName & myChar
	else 
		redim preserve nameList(tempNameCount)
		nameList(tempNameCount)= trim(tempName)
		tempNameCount = tempNameCount + 1
		tempName=""
	end if 
next
redim preserve nameList(tempNameCount)
nameList(tempNameCount)= trim(tempName)
'finish reading in name

textLine = objFile.ReadLine 'find initial month 
for i = 0 to 11
	if(instr(textline, months(i))>0) then
		previous = months(i)
		i = 11
	end if
next
for i = 2005 to 2015 'find initial year
	if(instr(textline, "" & i)) then
		currentYear = i
		i = 2015
	end if
next


redim monthlyMessageCount(ubound(nameList),0) 'array of each persons total message count
redim fileNames(0) 'year-month
redim fileContent(0) 'all messages and everything for each month

totalMonths = 0 'counts how many different months are gone through
output = ""
do while objFile.AtEndOfStream <> true
    
	
	infoLine = 0
	for i = 0 to ubound(nameList)
		if(instr(textLine, nameList(i)) = 1) then
			monthlyMessageCount(i, totalMonths) = monthlyMessageCount(i, totalMonths) + 1 'increment appropriate counters
			infoLine = 1
		end if
	next
	
	if(infoLine = 1) then
		
		'set myMonth to current month
		for i = 0 to 11 
			if(instr(textline, months(i))>0) then
				currentMonth = months(i)
				i = 12
			end if
		next
		
		if(currentMonth <> previous) then
			
			'put total count of messages at the end of content
			tempOutputAddon = ""
			for i = 0 to ubound(nameList)
				tempOutputAddon = tempOutputAddon & nameList(i) & ": " & monthlyMessageCount(i, totalMonths) & vbcrlf
			next
			output = tempOutputAddon & vbcrlf & output
			
			'add to content array
			redim preserve fileContent(totalMonths)
			fileContent(totalMonths) = output 
			
			'create title and place in title array
			redim preserve fileNames(totalMonths)
			fileNames(totalMonths) = currentYear & "-" & monthToNum(previous) & ".txt"
			
			'set myYear to current year
			for i = currentYear to 2015
				if(instr(textline, "" & i)>0) then
					currentYear = i
					i = 2015
				end if
			next
			
			'reset varaiables
			previous = currentMonth
			output = ""
			
			totalMonths = totalMonths + 1
			redim preserve monthlyMessageCount(ubound(nameList), totalMonths)
			
		end if
		
	end if
	output = output & textline & vbcrlf
	textLine = objFile.ReadLine
loop
objFile.Close
'put total count of messages at the end of content
tempOutputAddon = ""
for i = 0 to ubound(nameList)
	tempOutputAddon = tempOutputAddon & nameList(i) & ": " & monthlyMessageCount(i, totalMonths) & vbcrlf
next
output = tempOutputAddon & vbcrlf & output

'add to content array
redim preserve fileContent(totalMonths)
fileContent(totalMonths) = output 

'create title and place in title array
redim preserve fileNames(totalMonths)
fileNames(totalMonths) = currentYear & "-" & monthToNum(previous) & ".txt"

'set myYear to current year
for i = currentYear to 2015
	if(instr(textline, "" & i)>0) then
		currentYear = i
		i = 2015
	end if
next

'reset varaiables
previous = currentMonth
output = ""




wscript.echo "Start File write" 

'WRITE TO FILES
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateFolder(totalFileName & " files")
Set f = fso.CreateFolder(totalFileName & " files\seperated_month_files")

for i = 0 to totalMonths
	fileName = fileNames(i)
	Set objFile = objFSO.CreateTextFile(totalFileName & " files\seperated_month_files\" & fileName, true)  'Overwrites previous file
	objFile.WriteLine (fileContent(i))
	objFile.Close
next
wscript.echo "Finish file write"

Set objFile = objFSO.CreateTextFile(totalFileName & " files\totalStats.txt", true)  'create stat file with total counts

line = ""
for i = 0 to ubound(nameList)
	line = line & "	" & nameList(i) 
next
objFile.WriteLine (line)

for j = 0 to totalMonths
	line = left(fileNames(j), 7)
	for i = 0 to ubound(nameList)
		line = line & "	" & monthlyMessageCount(i, j) 
	next
	
	objFile.WriteLine (line)
next
objFile.Close