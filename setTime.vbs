Set HTTP = CreateObject("MSXML2.serverXMLHTTP")

HTTP.Open "GET", ("http://ts.sector572.com/DateTime/currentTime"), False

HTTP.setRequestHeader "Accept", "text/plain"

HTTP.send("")

WScript.StdOut.Write "Setting time to " & HTTP.responseText 

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.Run("cmd.exe /c time " & HTTP.responseText)