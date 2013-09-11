' Map Network Printers
' Connect to Shared Printers and set one as default
'
' by RaveMaker - http://ravemaker.net

Option Explicit

Dim objNetwork
Set objNetwork = CreateObject("Wscript.Network")

' Connect to printer called Floor2Left on PrinterServer
objNetwork.AddWindowsPrinterConnection "\\PrinterServer\Floor2Left"

' Set Floor2Left as default printer
objNetwork.SetDefaultPrinter "\\server\AdminPrinter"