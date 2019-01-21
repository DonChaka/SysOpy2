filename = "C:\Users\Bucek\Desktop\liczby.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(filename, 1, False, -1)

sum = 0
Do Until f.AtEndOfStream
  sum = sum + CLng(f.ReadLine)
Loop

WScript.Echo sum
f.Close