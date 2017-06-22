Dim FilePath, NrDays
FilePath = "C:\Proficy Historian Data\LogFiles"
NrDays = 7

Set objShell = CreateObject ("Shell.Application")
Set objFolder = objShell.Namespace (FilePath)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim arrHeader
Dim DateMod
arrHeader = objFolder.GetDetailsOf (objFolder.Items, 3)
For Each strFileName in objFolder.Items
 DateMod = objFolder.GetDetailsOf (strFileName, 3) 
 if datediff("d",DateMod,Now()) > NrDays and _
    Instr(1,strFileName,".log",1) > 0 _
 then
  on error resume next
  objFSO.DeleteFile FilePath & "\" & strFileName,True
  on error goto 0
 end if
 'Wscript.echo DateMod & " " & strFileName
Next

