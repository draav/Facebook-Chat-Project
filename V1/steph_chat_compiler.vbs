REM Nicholas Devlin
REM 08/09/2013
REM This code reads in message history from steph then creates a new file for each month
' NEW LINE = vbcrlf

'READ/ PROCESS***********

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("steph edited copy.txt", 1) 'Open file to read

previous = "May"
stephCount = 0
nickCount = 0
myYear = 2011
myMonth = 5
output = ""
dim fileNames(36)
dim fileContent(36)
dim sTotalcount(36)
dim nTotalCount(36)
counter = 0
do while objFile.AtEndOfStream <> true
    textLine = objFile.ReadLine
    name = mid(textLine,1,11)
	if(name = "Steph Dudak" OR name = "Nick Devlin") then
		mName = ""
		for i = 13 to len(textLine)
			myChar = Mid(textLine, i, 1)
			if(myChar <> " ") then
				mName = mName & myChar
			else 
				i = len(textLine)
			end if 
		next
		
		if(mName <> previous) then
			
			sTotalCount(counter)=stephCount
			nTotalCount(counter)=nickCount
			output = output & vbcrlf & "Steph: " & stephCount & vbcrlf & "Nick: " & nickCount
			
			fileContent(counter) = output
			if(myMonth < 10) then fileNames(counter) = myYear & "-0" & myMonth & ".txt"
			if(myMonth >= 10) then fileNames(counter) = myYear & "-" & myMonth & ".txt"
			myMonth = myMonth + 1
			if(mName="January") then 
				myYear = myYear + 1 
				myMonth = 1
			end if
			
			
			
			previous = mName
			output = ""
			stephCount = 0
			nickCount = 0
			counter = counter + 1
		end if
		
		if(name = "Steph Dudak") then stephCount = stephCount + 1
		if(name = "Nick Devlin") then nickCount = nickCount + 1
		
	end if
	output = output & textline & vbcrlf
	
loop
objFile.Close


wscript.echo "Start: " &fileName

REM 'WRITE*****************
REM filePath = "C:\Users\Owner\Documents\0 Nicholas' stuff\Facebook Chat Project\Steph\"

for j = 0 to counter-1
	fileName = fileNames(j)
	Set objFile = objFSO.CreateTextFile(fileName, true)  'Overwrites previous file
	objFile.WriteLine (fileContent(j))
	objFile.Close
next
wscript.echo "Success"

Set objFile = objFSO.CreateTextFile("totalStats.txt", true)  'Overwrites previous file

for j = 0 to counter-1
	line = (j+1)&". Steph:"&sTotalCount(j)&"	Nick: "&nTotalCount(j)
	objFile.WriteLine (line)
next
objFile.Close