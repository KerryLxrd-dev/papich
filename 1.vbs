Dim WShell
Set WShell = CreateObject("WScript.Shell")
WShell.Run "KEK.vmp.exe -f json", 0
Set WShell = Nothing
