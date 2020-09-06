Sub UnHideWorksheets()
    Dim sSheetName As String
    Dim w As Worksheet
    Dim sTemp As String

    sTemp = "Name (or partial) of sheet to show?"
    sSheetName = InputBox(sTemp, "Show Hidden Sheet")
    If sSheetName > "" Then
        sSheetName = LCase(sSheetName)
        For Each w In Sheets
            w.Tab.ColorIndex = xlColorIndexNone
            sTemp = LCase(w.Name)
            If InStr(sTemp, sSheetName) Then
                w.Visible = True
                w.Tab.ColorIndex = 6
            End If
        Next w
    End If
End Sub
