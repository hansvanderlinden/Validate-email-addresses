'--------------------------------------------------------------------------
'- Script: verifyEmail.vbs
'- Usage: verify email addresses from excel spreadsheet
'- Parameters:
'		xls: path and name of spreadheet containing email addresses
'- Returns: nothing
'- Output: result of verification in 12th column, uniqueness in 13th column
'- Date created: 22-10-2008
'- Date changed:
'--------------------------------------------------------------------------

'- define cols
Const emailCol = 6
Const errorCol = 7
Const uniqueCol = 8

Set re = New RegExp
re.IgnoreCase = True
re.Global = True

xls = "G:\Marketing Communicatie & Sales Support\WebSite and apps\_Tools\Validate email address\emails.xlsx"
call Main(xls)
Set re = nothing

Sub Main(byval xls)
	'- craete excel object
	Set objExcel = CreateObject("Excel.Application")
	
	'- open sheet
	On Error Resume Next
	objExcel.WorkBooks.Open xls
	If Err.Number <> 0 Then
		WScript.Echo "Error opening " & xls & ". Could not find the file."
		WScript.Quit
	End If
	On Error Goto 0
	
	Set emails = objExcel.ActiveWorkbook.Worksheets(1)

	'- set headers for errors and uniqueness
	emails.Cells(1, errorCol).Value = "Error"
	emails.Cells(1, uniqueCol).Value = "Uniqueness"

	intRow = 2
	unique = ""
	Do While emails.Cells(intRow, emailCol).Value <> ""
		mailerror = "valid"
		email = Trim(emails.Cells(intRow, emailCol).Value)
	    
	    If InStr(unique, email) = 0 Then
			'check syntax
			re.Pattern= "^[-!#$%&\'*+\\./0-9=?A-Z^_`a-z{|}~]+@[-!#$%&\'*+\\/0-9=?A-Z^_`a-z{|}~]+\.[-!#$%&\'*+\\./0-9=?A-Z^_`a-z{|}~]+$"
			If Not re.Test(email) Then 	mailerror = "Incorrect syntax"
					
			If mailerror = "valid" Then 'check MX record
				'extract domain
				domain = Mid(email,InStr(email, "@")+1)
			
				Set oShell = WScript.CreateObject("WScript.Shell")
				Set fso = WScript.CreateObject("Scripting.FileSystemObject")
				sTemp = oShell.ExpandEnvironmentStrings("%TEMP%")
				sTempFile = sTemp & "\" & fso.GetTempName
				
				cmd = "%comspec% /c nslookup -type=MX " & domain & ">" & sTempFile
				oShell.Run cmd, 0, True
				
				Set fFile = fso.OpenTextFile(sTempFile, 1, 0, -2)
				sResults = fFile.ReadAll
				fFile.Close
				fso.DeleteFile(sTempFile)
				Set oShell = Nothing
				Set fso = Nothing
				
				' check for DNS or MX errors
				re.Pattern= "\n"
				Set matches = re.Execute(sResults)
				If matches.count <= 3 Then
					mailerror = "invalid; no MX"
				Else
					re.Pattern = "MX"
					Set matches = re.Execute(sResults)
					If matches.count = 0 Then mailerror = "invalid; no exchange"
				End If
		    End If
		    
			emails.Cells(intRow, errorCol).Value = mailerror
			emails.Cells(intRow, uniqueCol).Value = "unique"
			
			If mailerror <> "valid" Then
				WScript.echo sResults
				WScript.Echo intRow & " - " & email & ": " & mailerror & " - unique"
			End If
			unique = unique & "|" & emails.Cells(intRow, emailCol).Value
		Else
			emails.Cells(intRow, uniqueCol).Value = "double"
			WScript.Echo intRow & " - " & email & ": " & " - double"
		End If

		intRow = intRow + 1
	Loop
	
	'- Save
	objExcel.ActiveWorkbook.Save
	
	' Close workbook and quit Excel.
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit
	
	' Clean up.
	Set objExcel = Nothing
	Set emails = Nothing
End Sub