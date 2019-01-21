Function ShowFolderList(folderspec)
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFolder(folderspec)
   Set fc = f.SubFolders
   For Each f1 in fc
	  folder = folderspec & "\" & f1.name
	  s = s & showfolderlist(folder)
      s = s & f1.name 
      s = s & vbNewLine
   Next
   ShowFolderList = s
End Function

wscript.echo ShowFolderList("h:\pobrane\SO2-skrypty")