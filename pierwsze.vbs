Function czy_pierwsza(liczba)
 czy_pierwsza = true
 For i = 2 To liczba-1
  If ( liczba mod i = 0 ) Then
   czy_pierwsza = false
   Exit Function
  End If
 Next
End Function

wynik = "Liczby pierwsze: "

For j = 2 To 100
 If (czy_pierwsza(j)) Then
  tmp = CStr(j)
  wynik = wynik + tmp +	 " "
 End If
Next

MsgBox wynik