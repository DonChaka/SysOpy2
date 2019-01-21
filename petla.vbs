Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fso, MyFile, FileName, TextLine

Set fso = CreateObject("Scripting.FileSystemObject")

FileName = "e:\Pulpit\skrypty\test.txt"

Set MyFile = fso.OpenTextFile(FileName, ForWriting, True)

MyFile.WriteLine "1"
MyFile.WriteLine "2"
MyFile.WriteLine "3"
MyFile.WriteLine "4"
MyFile.WriteLine "5"
MyFile.Close

Set MyFile = fso.OpenTextFile(FileName, ForReading)

Do While MyFile.AtEndOfStream <> True
    TextLine = MyFile.ReadLine
    temp = cint(textline)
    wynik = wynik + temp
Loop

wscript.echo wynik

MyFile.Close