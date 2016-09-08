Set HTTP = CreateObject("MSXML2.serverXMLHTTP")

HTTP.Open "GET", ("https://ts.sector572.com/DateTime/windows/currentDate"), False

HTTP.setRequestHeader "Accept", "text/plain"

HTTP.send("")

Dim date 
date = HTTP.responseText 

Set regex = New RegExp 
regex.Pattern = "^[0-9]{2}/[0-9]{2}/[0-9]{4}$" 

If regex.Test(date) Then
	WScript.StdOut.WriteLine "Setting date to " & date 
	

	Set WshShell = WScript.CreateObject("WScript.Shell")

	WshShell.Run("cmd.exe /c date " & date) 
Else 
	WScript.StdOut.WriteLine "Invalid date format received." 
End If 

regex.Pattern = "^[0-9]{2}:[0-9]{2}:[0-9]{2}$" 

HTTP.Open "GET", ("https://ts.sector572.com/DateTime/windows/currentTime"), False

HTTP.setRequestHeader "Accept", "text/plain"

HTTP.send("")

Dim time 
time = HTTP.responseText 

If regex.Test(time) Then
	WScript.StdOut.WriteLine "Setting time to " & time 

	Set WshShell = WScript.CreateObject("WScript.Shell")

	WshShell.Run("cmd.exe /c time " & time)
Else 
	WScript.StdOut.WriteLine "Invalid time format received." 
End If 
