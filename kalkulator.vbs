Do While true
 pierwsza = InputBox("Podaj pierwsza liczbe (lub 'koniec' by zakonczyc): ")
 If pierwsza = "koniec" Then
	Exit Do
 End If
 druga = InputBox("Podaj druga liczbe: ")
 operacja = InputBox("Podaj operacje: ")

 If (operacja = "+") Then
  wynik = "Wynik dodawania: " & CInt(pierwsza) + CInt(druga)
 ElseIf (operacja = "-") Then
  wynik = "Wynik odejmowania: " & CInt(pierwsza) - CInt(druga)
 ElseIf (operacja = "*") Then
  wynik = "Wynik mnozenia: " & CInt(pierwsza) * CInt(druga)
 ElseIf (operacja = "/") Then
  If (druga = "0") Then
   wynik = "Nie mozna dzielic przez zero!"
  Else
   wynik = "Wynik dzielenia: " & CInt(pierwsza) / CInt(druga)
  End If
 Else
  wynik = "Bledna operacja!"
 End If
 MsgBox wynik
Loop

MsgBox "Koniec programu, dzieki za uzycie :)"