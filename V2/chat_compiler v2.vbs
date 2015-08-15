REM Nicholas Devlin
REM 08/13/2013
REM This code reads in message history then creates a new file for each month
REM Changes: allow months to be skipped in chat, allow days of week to be included, add functions & code to increase reusablility for other people
' NEW LINE = vbcrlf

wscript.echo "Begin Code: "

months = array("January","February","March","April","May","June","July","August","September","October","November","December")

'FUNCTIONS
function nextWord(givenString, startPoint, wordCount)
nextWord = ""
for i = startPoint to len(givenString)
	myChar = Mid(textLine, i, 1)
	if(myChar <> " " and myChar <> "," and wordCount <> 0) then
		nextWord = nextWord & myChar
		wordCount = wordCount - 1
	else 
		i = len(givenString)
	end if 
next
end function

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
	monthToNum = -1
end select

end function


'READ/ PROCESS
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateFolder("seperate_month_files")

set objFSO = CreateObject("Scripting.FileSystemObject")
set objFile = objFSO.OpenTextFile("steph full chat.txt", 1) 'Open file to read

'variables that change per file, should add code before the loop that grabs the info from file instead of hard coding it in every time
previous = "May"
myName = "Nick Devlin"
friendName = "Steph Dudak"
currentYear = 2011

myCount = 0
friendCount = 0

output = ""

dim fileNames(36) 'year-month
dim fileContent(36)
dim myMessageCount(36)
dim friendMessageCount(36)
counter = 0 'counts how many different months are gone through

textLine = objFile.ReadLine


do while objFile.AtEndOfStream <> true
    textLine = objFile.ReadLine
	
	if(left(textLine, len(myName)) = myName OR left(textLine, len(friendName)) = friendName) then
		
		'set myMonth
		for i = 0 to 11
			if(instr(textline, months(i))>0) then
				currentMonth = months(i)
				i = 12
			end if
		next
		
		if(currentMonth <> previous) then
			
			
			'set myYear
			for i = currentYear to 2015
				if(instr(textline, "" & i)) then
					currentYear = i
					i = 2015
				end if
			next
			
			'place counters to respective arrays
			friendMessageCount(counter)=friendCount
			myMessageCount(counter)=myCount
			
			'put total count of messages at the end of content
			output = output & vbcrlf & friendName& ": " & friendCount & vbcrlf & myName & ": " & myCount
			
			'add to content array
			fileContent(counter) = output 
			
			'create title and place in title array
			fileNames(counter) = currentYear & "-" & monthToNum(previous) & ".txt"
			'reset varaiables
			previous = currentMonth
			output = ""
			friendCount = 0
			myCount = 0
			
			counter = counter + 1
		end if
		
		if(left(textLine, len(friendName)) = friendName) then friendCount = friendCount + 1
		if(left(textLine, len(myName)) = myName) then myCount = myCount + 1
		
	end if
	output = output & textline & vbcrlf
	
loop
objFile.Close


wscript.echo "Start File write: " 

'WRITE TO FILES

for j = 0 to counter-1
	fileName = fileNames(j)
	Set objFile = objFSO.CreateTextFile("seperate_month_files\" & fileName, true)  'Overwrites previous file
	objFile.WriteLine (fileContent(j))
	objFile.Close
next
wscript.echo "Finish file write"

Set objFile = objFSO.CreateTextFile("totalStats.txt", true)  'create stat file with total counts
objFile.WriteLine (friendName&"	"&myName)
for j = 0 to counter-1
	line = friendMessageCount(j)&"	"&myMessageCount(j)
	objFile.WriteLine (line)
next
objFile.Close