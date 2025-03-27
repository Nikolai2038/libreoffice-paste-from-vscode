' the actual macro code For copy&paste including indentation,
' it needs the space characters Set in VSCode As indentation,
' Not tabs (in Case of tabs Set ASCII code needs To be modified in the script)
'  the macro can be assigned To a shorcut, For example  ex. ctrl+D

Sub Kopiowanie2 ' the macro code

rem Get access To the document:
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")


Dim tekst As string
'Dim tekst2 As string
Dim tablicaLini(1000) As string
Dim liczbaWierszy As integer
Dim liniaTekstu As string
'Dim liczbaSpacji As integer
Dim lancuchSpacji As string

tekst = getClipboardText()

' proby z kursorem, tj wczytywanie pozycji, docelowo dodanie spacji
' chyba nawet nie potrzebne

' => wklejenie tekstu ze schowka z CSS

dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())

' => wczytanie liczby lini wklejonego kodu

liczbaWierszy = zliczWiersze(tekst)

' => przesuniecie kursora Do samej góry wklejonego tekstu

Dim cursor As Object

cursor = ThisComponent.getCurrentController.getViewCursor
cursor.gotoStartOfLine(False)

cursor.goUp(liczbaWierszy + 1, False)

' => wczytanie Do tablicy wszystkich linii tekstu

tablicaLini = wczytajWszystkieWiersze(tekst)

' => deklaracja obiektu Do wydruku tekstu

Dim args2(0) As New com.sun.star.beans.PropertyValue
args2(0).Name = "Text"

' => start pętli

Dim I As integer
Dim Z As integer

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

    cursor.goDown(1, False)
Next I ' koniec pętli

end Sub

Function zczytajSpacje(tekst As String) As string

    Dim zczytaneSpacje As string
    zczytaneSpacje = ""

    Dim Z As integer
    Z = Len(tekst)

    Dim I As integer
    Dim str As string

    For I = 1 To Z Step 1
    str = Mid(tekst,I,1)

    If Asc(str) = 32 Then
    zczytaneSpacje = zczytaneSpacje + " "
    Else
    Exit For
    End If


    zczytajSpacje = zczytaneSpacje

End Function

Function zliczSpacje(tekst As String) As Integer

    ' przeiterowanie
    Dim liczbaSpacji As Integer
    liczbaSpacji = 0

    Dim Z As integer
    Z = Len(tekst)

    Dim I As integer
    Dim str As string

    For I = 1 To Z Step 1
        str = Mid(tekst,I,1)

        If Asc(str) = 32 Then
            liczbaSpacji = liczbaSpacji + 1
        Else
            Exit For
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

    Dim tekst2 As string
    tekst2 = ""

    Dim liczbaWierszy As integer
    liczbaWierszy = 1

    Dim str As string
    Dim wiersz As string

    tekst = getClipboardText()

    Dim Z As integer
    Z = Len(tekst)

    Dim I As integer

    For I = 1 To Z Step 1
        str = Mid(tekst,I,1)

        wszystkieWiersze(liczbaWierszy) = wszystkieWiersze(liczbaWierszy) + str

        'If numerWiersza = liczbaWierszy Then
        '	   wiersz = wiersz + str
        'End If

        If Asc(str) = 10 Then ' i To dziala dobrze, tj. zwraca OSTATNI wiersz
            'wiersz = ""
            liczbaWierszy = liczbaWierszy + 1
        End If

    Next I

    wczytajWszystkieWiersze = wszystkieWiersze()

End Function ' ~ wczytajWiersz
'===================================================
'===================================================

' sprawna FUNKCJA ===================
Function wczytajWiersz(tekst As String, numerWiersza As Integer) As String

    Dim tekst2 As string
    tekst2 = ""

    Dim liczbaWierszy As integer
    liczbaWierszy = 1

    Dim str As string
    Dim wiersz As string

    tekst = getClipboardText()

    Dim Z As integer
    Z = Len(tekst)

    Dim I As integer

    For I = 1 To Z Step 1
        str = Mid(tekst,I,1)

        If numerWiersza = liczbaWierszy Then
            wiersz = wiersz + str
        End If

        If Asc(str) = 10 Then ' i To dziala dobrze, tj. zwraca OSTATNI wiersz
            'wiersz = ""
            liczbaWierszy = liczbaWierszy + 1
        End If

    Next I

    wczytajWiersz = wiersz

End Function ' ~ wczytajWiersz


' sprawna FUNKCJA ===================
Function zliczWiersze(tekst As String) As Integer

    Dim tekst2 As string
    tekst2 = ""

    Dim liczbaWierszy As integer
    liczbaWierszy = 0

    Dim str As string

    tekst = getClipboardText()

    Dim Z As integer
    Z = Len(tekst)

    Dim I As integer

    For I = 1 To Z Step 1
        str = Mid(tekst,I,1)

        If Asc(str) = 10 Then
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