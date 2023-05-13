' Deletes settings & cache
Dim path : path = SDB.ApplicationPath&"Scripts\Auto\"
Dim i : i = InStrRev(SDB.Database.Path,"\")
Dim appPath : appPath = Left(SDB.Database.Path,i)
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")


iniSec = "FilterNodes_MM"&Round(SDB.VersionHi)     'Put ini section name here
SDB.IniFile.DeleteSection(iniSec)
iniSec = "FilterNodes-Nodes_MM"&Round(SDB.VersionHi)     'Put ini section name here
SDB.IniFile.DeleteSection(iniSec)

If fso.FileExists(path&"FilterNodes.vbs") Then
	Call fso.DeleteFile(path&"FilterNodes.vbs")
End If

MsgBox("I hope your experiences with Filter Nodes were not all bad." & vbNewLine & "Please restart MediaMonkey for the uninstall to complete.")




