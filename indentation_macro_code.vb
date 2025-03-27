' The actual macro code for copy & paste including indentation.
' It needs space characters set in VSCode as indentation, not tabs.
' If tabs are used, the ASCII code needs to be modified in the script.
' The macro can be assigned to a shortcut, for example, Ctrl+D.

Sub CopyIndentation ' The macro code

    ' Get access to the document:
    Dim document As Object
    Dim dispatcher As Object
    document = ThisComponent.CurrentController.Frame
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    Dim text As String
    Dim linesArray() As String
    Dim lineCount As Integer
    Dim textLine As String
    Dim spaceString As String

    text = GetClipboardText()

    ' Paste text from clipboard
    dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())

    ' Count the number of lines in the pasted text
    lineCount = CountLines(text)

    ' Move cursor to the top of the pasted text
    Dim cursor As Object
    cursor = ThisComponent.getCurrentController.getViewCursor()
    cursor.gotoStartOfLine(False)
    cursor.goUp(lineCount, False)

    ' Read all lines into an array
    linesArray = ReadAllLines(text)

    ' Declare object for text insertion
    Dim args2(0) As New com.sun.star.beans.PropertyValue
    args2(0).Name = "Text"

    ' Start loop
    Dim i As Integer
    For i = 0 To lineCount - 1
        ' Get the i-th line of text
        textLine = linesArray(i)

        ' Extract leading spaces from the line
        spaceString = ExtractSpaces(textLine)

        ' Move cursor to the beginning of the line
        cursor.gotoStartOfLine(False)

        ' Insert the corresponding number of spaces
        args2(0).Value = spaceString
        dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args2())

        ' Move cursor to the next line
        cursor.goDown(1, False)
    Next i ' End loop

End Sub

Function ExtractSpaces(text As String) As String
    Dim extractedSpaces As String
    extractedSpaces = ""

    Dim i As Integer
    Dim str As String

    For i = 1 To Len(text)
        str = Mid(text, i, 1)

        If Asc(str) = 32 Then
            extractedSpaces = extractedSpaces & " "
        Else
            Exit For
        End If
    Next i

    ExtractSpaces = extractedSpaces
End Function

Function CountLines(text As String) As Integer
    Dim lineCount As Integer
    lineCount = 1

    Dim i As Integer
    For i = 1 To Len(text)
        If Asc(Mid(text, i, 1)) = 10 Then
            lineCount = lineCount + 1
        End If
    Next i

    CountLines = lineCount
End Function

Function ReadAllLines(text As String) As Variant
    Dim allLines() As String
    Dim lineIndex As Integer
    Dim i As Integer
    Dim str As String
    Dim currentLine As String

    lineIndex = 0
    currentLine = ""

    For i = 1 To Len(text)
        str = Mid(text, i, 1)
        If Asc(str) = 10 Then
            ReDim Preserve allLines(lineIndex)
            allLines(lineIndex) = currentLine
            lineIndex = lineIndex + 1
            currentLine = ""
        Else
            currentLine = currentLine & str
        End If
    Next i

    If currentLine <> "" Then
        ReDim Preserve allLines(lineIndex)
        allLines(lineIndex) = currentLine
    End If

    ReadAllLines = allLines
End Function

Function GetClipboardText() As String ' Retrieves plain text from clipboard
    Dim oClip As Object ' com.sun.star.datatransfer.clipboard.SystemClipboard
    Dim oConverter As Object ' com.sun.star.script.Converter
    Dim oClipContents As Object
    Dim oTypes As Object
    Dim i As Integer
    Dim result As String

    oClip = createUnoService("com.sun.star.datatransfer.clipboard.SystemClipboard")
    oConverter = createUnoService("com.sun.star.script.Converter")
    On Error Resume Next
    oClipContents = oClip.getContents()
    oTypes = oClipContents.getTransferDataFlavors()

    result = ""

    For i = LBound(oTypes) To UBound(oTypes)
        If oTypes(i).MimeType = "text/plain;charset=utf-16" Then
            result = oConverter.convertToSimpleType(oClipContents.getTransferData(oTypes(i)), com.sun.star.uno.TypeClass.STRING)
            Exit For
        End If
    Next i

    GetClipboardText = result
End Function
