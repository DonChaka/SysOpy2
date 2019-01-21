path = inputbox("Podaj przeszukiwana sciezke")
path = "H:\Pobrane\SO2-skrypty"

wybor = inputbox ("wybierz kryterium:" & vbnewline &_
"1 - dany rozmiar lub mniejsze" & vbnewline &_
"2 - pierwsza litera" & vbnewline &_
"3 - rozszerzenie")

zakres = inputbox("Podaj kryterium: ")

if wybor = "1" then
	set fso = CreateObject("Scripting.FileSystemObject")
	set folder = fso.GetFolder(path)
	set file = folder.files
	for each temp in file
		if temp.size <= CInt(zakres) then
			wscript.echo temp
		end if
	next
elseif wybor = "2" then
	set fso = CreateObject("Scripting.FileSystemObject")
	set folder = fso.GetFolder(path)
	set file = folder.files
	for each temp in file
		if left(temp.name, 1) = zakres then
			wscript.echo temp
		end if
	next
elseif wybor = "3" then
	set fso = CreateObject("Scripting.FileSystemObject")
	set folder = fso.GetFolder(path)
	set file = folder.files
	for each temp in file
		if fso.GetExtensionName(temp) = zakres then
			wscript.echo temp
		end if
	next
else
	msgbox "Bledne dane"
end if