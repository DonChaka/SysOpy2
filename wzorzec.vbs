
Path = "E:\Pulpit\skrypty"
wzorzec = CStr( InputBox("Podaj wzorzec: ") )
Set oFSO = CreateObject("Scripting.FileSystemObject") 
Set folder = oFSO.GetFolder(Path)
 
WScript.Echo (Path+"\"+wzorzec)
If oFSO.FileExists(Path+"\"+wzorzec) Then
  WScript.Echo Path+"\"+wzorzec
  Else
  WScript.Echo "Brak plikow pasujacych do wzorca"
End If

Set subfolders = folder.subfolders 
MsgBox subfolders.Count