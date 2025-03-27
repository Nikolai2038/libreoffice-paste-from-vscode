' the actual macro code for copy&paste including indentation,
' it needs the space characters set in VSCode as indentation,
 ' not tabs (in case of tabs set ASCII code needs to be modified in the script)
'  the macro can be assigned to a shorcut, for example  ex. ctrl+D

sub Kopiowanie2 ' the macro code

rem get access to the document:
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")


Dim tekst as string
'Dim tekst2 as string
Dim tablicaLini(1000) as string
Dim liczbaWierszy as integer
Dim liniaTekstu as string
'Dim liczbaSpacji as integer
Dim lancuchSpacji as string

tekst = getClipboardText()

' proby z kursorem, tj wczytywanie pozycji, docelowo dodanie spacji
' chyba nawet nie potrzebne

' => wklejenie tekstu ze schowka z CSS

	dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())

' => wczytanie liczby lini wklejonego kodu

	liczbaWierszy = zliczWiersze(tekst)

' => przesuniecie kursora do samej góry wklejonego tekstu

	Dim cursor As Object

    cursor = ThisComponent.getCurrentController.getViewCursor
    cursor.gotoStartOfLine(false)

    cursor.goUp(liczbaWierszy + 1, false)

' => wczytanie do tablicy wszystkich linii tekstu

	tablicaLini = wczytajWszystkieWiersze(tekst)

' => deklaracja obiektu do wydruku tekstu

	dim args2(0) as new com.sun.star.beans.PropertyValue
	args2(0).Name = "Text"

' => start pętli

	Dim I as integer
	Dim Z as integer

	Z = liczbaWierszy + 1

For I = 1 To Z Step 1

' => pobranie i-tej linii tekstu

 	liniaTekstu = tablicaLini(I)

' => wczytanie łańcucha wiodących spacji z liniTekstu

    lancuchSpacji = zczytajSpacje(liniaTekstu)

' => wydruk odpowiedniej liczby spacji

	args2(0).Value = lancuchSpacji

	dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args2())

	' ruch kursorem

	cursor.goDown(1, false)
Next I ' koniec pętli

end sub

Function zczytajSpacje(tekst As String) As string

	Dim zczytaneSpacje As string
	zczytaneSpacje = ""

	Dim Z as integer
	Z = Len(tekst)

	Dim I as integer
	Dim str as string

	For I = 1 To Z Step 1
   		str = Mid(tekst,I,1)

   	if Asc(str) = 32 Then
   	  zczytaneSpacje = zczytaneSpacje + " "
   	Else
   		exit for
   	End If
Next I

	zczytajSpacje = zczytaneSpacje

End Function

Function zliczSpacje(tekst As String) As Integer

   ' przeiterowanie
Dim liczbaSpacji As Integer
liczbaSpacji = 0

Dim Z as integer
Z = Len(tekst)

Dim I as integer
Dim str as string

For I = 1 To Z Step 1
   str = Mid(tekst,I,1)

   if Asc(str) = 32 Then
   	  liczbaSpacji = liczbaSpacji + 1
   Else
   	exit for
   End If

Next I

	zliczSpacje = liczbaSpacji
End Function

'---------------------------------------------------
'	~zliczanie spacji
'---------------------------------------------------


'===================================================
' sprawna funkcja
'===================================================
Function wczytajWszystkieWiersze(tekst As String)

Dim wszystkieWiersze(1000) As String

Dim tekst2 as string
tekst2 = ""

Dim liczbaWierszy as integer
liczbaWierszy = 1

Dim str as string
Dim wiersz as string

tekst = getClipboardText()

Dim Z as integer
Z = Len(tekst)

Dim I as integer

For I = 1 To Z Step 1
   str = Mid(tekst,I,1)

   wszystkieWiersze(liczbaWierszy) = wszystkieWiersze(liczbaWierszy) + str

   'if numerWiersza = liczbaWierszy Then
   '	   wiersz = wiersz + str
   'End If

   if Asc(str) = 10 Then ' i to dziala dobrze, tj. zwraca OSTATNI wiersz
   	  'wiersz = ""
   	  liczbaWierszy = liczbaWierszy + 1
   End If

Next I

	wczytajWszystkieWiersze = wszystkieWiersze()

End Function ' ~ wczytajWiersz
'===================================================
'===================================================

' sprawna FUNKCJA ===================
Function wczytajWiersz(tekst As String, numerWiersza as Integer) As String

Dim tekst2 as string
tekst2 = ""

Dim liczbaWierszy as integer
liczbaWierszy = 1

Dim str as string
Dim wiersz as string

tekst = getClipboardText()

Dim Z as integer
Z = Len(tekst)

Dim I as integer

For I = 1 To Z Step 1
   str = Mid(tekst,I,1)

   if numerWiersza = liczbaWierszy Then
	   wiersz = wiersz + str
   End If

   if Asc(str) = 10 Then ' i to dziala dobrze, tj. zwraca OSTATNI wiersz
   	  'wiersz = ""
   	  liczbaWierszy = liczbaWierszy + 1
   End If

Next I

	wczytajWiersz = wiersz

End Function ' ~ wczytajWiersz


' sprawna FUNKCJA ===================
Function zliczWiersze(tekst As String) As Integer

Dim tekst2 as string
tekst2 = ""

Dim liczbaWierszy as integer
liczbaWierszy = 0

Dim str as string

tekst = getClipboardText()

Dim Z as integer
Z = Len(tekst)

Dim I as integer

For I = 1 To Z Step 1
   str = Mid(tekst,I,1)

   if Asc(str) = 10 Then
   	  tekst2 = tekst2 + "L"
   	  liczbaWierszy = liczbaWierszy + 1
   End If

Next I

	zliczWiersze = liczbaWierszy

End Function ' ~ zliczWiersze

Function getClipboardText() As String rem +++ sprawna funkcja, zwraca CZYSTY tekst, z tabulacją
    '''Returns a string of the current clipboard text'''

    Dim oClip As Object ' com.sun.star.datatransfer.clipboard.SystemClipboard
    Dim oConverter As Object ' com.sun.star.script.Converter
    Dim oClipContents As Object
    Dim oTypes As Object
    Dim i%

    oClip = createUnoService("com.sun.star.datatransfer.clipboard.SystemClipboard")
    oConverter = createUnoService("com.sun.star.script.Converter")
    On Error Resume Next
    oClipContents = oClip.getContents
    oTypes = oClipContents.getTransferDataFlavors

    For i = LBound(oTypes) To UBound(oTypes)
        If oTypes(i).MimeType = "text/plain;charset=utf-16" Then
            Exit For
        End If
    Next

    If (i >= 0) Then
        On Error Resume Next
        getClipboardText = oConverter.convertToSimpleType _
            (oClipContents.getTransferData(oTypes(i)), com.sun.star.uno.TypeClass.STRING)
    End If

End Function ' getClipboardText