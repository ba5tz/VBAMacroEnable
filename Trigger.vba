Option Explicit
Dim xlApp, xlBook
Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Open("C:\lokasi\File\MacroTest.xlsm", 0, True)
xlApp.Run "MacroEnabled"
xlBook.Close
xlApp.Quit
Set xlBook = Nothing
Set xlApp = Nothing
WScript.Quit
