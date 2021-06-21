' SearchTBox.bas
' Finds, highlights, and counts all instances of a matching string within text boxes
' Written by John Gabatin

Attribute VB_Name = "Module1"

Private Sub SearchTBox()
    Dim startIndex As Range
    Dim tBox As TextBox
    Dim tFind As String
    Dim sTemp As String
    Dim countMatch As Integer
    Dim notFound As Boolean

    ' String to find
    tFind = InputBox("Search for:")

    If Trim(tFind) = "" Then
        MsgBox "Empty string entered."
    Exit Sub
    End If

    Set startIndex = ActiveCell
    notFound = True

    ' Traverse over all TextBoxes
    For Each tBox In ActiveSheet.TextBoxes
        sTemp = tBox.Text

        ' Found the matching string (not case sensitive)
        If InStr(LCase(sTemp), LCase(tFind)) <> 0 Then
            ' Keep a running count
            countMatch = countMatch + 1
            tBox.Select

            ' Color string white and TextBox black
            With tBox.Characters(Start:=iPos, Length:=Len(tFind)).Font
                .ColorIndex = 2
                .Bold = True
                Selection.ShapeRange.Fill.ForeColor.RGB = RGB(0, 0, 0)
            End With

            notFound = False
        End If
    Next

  If notFound Then
    MsgBox "No matches found."
  Exit Sub
  End If

  MsgBox "Matches found: " & countMatch, vbInformation
  startIndex.Select
  Set startIndex = Nothing

End Sub
