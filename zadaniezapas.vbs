function f_rozmiar(path, zakres)
	set fso = CreateObject("Scripting.FileSystemObject")
	set folder = fso.GetFolder(path)
	set file = folder.files
	for each temp in file
		if temp.size <= CInt(zakres) then
			wscript.echo temp.name
		end if
	next
end function

function f_litery(path, zakres)
	set fso = CreateObject("Scripting.FileSystemObject")
	set folder = fso.GetFolder(path)
	set file = folder.files
	for each temp in file
		if right(fso.getfilename(file), 1) = zakres then
			wscript.echo temp.name
		end if
	next
end function

function f_rozszerzenie(path, zakres)
	set fso = CreateObject("Scripting.FileSystemObject")
	set folder = fso.GetFolder(path)
	set file = folder.files
	for each temp in file
		if fso.getextensionname(file) = zakres then
			wscript.echo temp.name
		end if
	next
end function


path = inputbox("Podaj przeszukiwana sciezke")

wybor = inputbox ("wybierz kryterium:" & vbnewline &_
"1 - dany rozmiar lub mniejsze" & vbnewline &_
"2 - pierwsza litera" & vbnewline &_
"3 - rozszerzenie")

zakres = inputbox("Podaj kryterium: ")

if wybor = "1" then
	f_rozmiar(path, zakres)
else if wybor = "2" then
	f_litery(path, zakres)
else if wybor = "3" then
	f_rozszerzenie(path,zakres)
else
	msgbox "Bledne dane"
end if




