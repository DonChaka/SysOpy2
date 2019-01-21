Function ShowFolderList(folderspec)
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFolder(folderspec)
   Set fc = f.Files
   For Each f1 in fc
	'if fso.GetExtensionName(f1)  = "txt" then
		wscript.echo fso.GetFileName(f1)
	'end if
   Next
   ShowFolderList = s
End Function

showfolderlist("h:\pobrane\SO2-skrypty")