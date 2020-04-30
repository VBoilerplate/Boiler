Public Sub CopyWorksheet(wksName As String)
    
    Dim newName As String
    newName = wksName & "_w"
    
    If WorksheetNameIsPresent(newName) Then
        Application.DisplayAlerts = False
        Worksheets(newName).Delete
        Application.DisplayAlerts = True
    End If
    
    Dim wks As Worksheet
    Dim newWks As Worksheet
    
    Set wks = Worksheets(wksName)
    wks.Copy after:=Worksheets(Worksheets.Count)
    Set newWks = Worksheets.Item(Worksheets.Count)
    
    With newWks
        .Name = newName
        .Tab.Color = vbBlue
    End With
    
End Sub

Public Function WorksheetNameIsPresent(newName As String) As Boolean

    Dim wks As Worksheet
    For Each wks In ThisWorkbook.Worksheets
        If wks.Name = newName Then
            WorksheetNameIsPresent = True
            Exit Function
        End If
    Next wks
    WorksheetNameIsPresent = False
    
End Function

Public Sub CopyWorksheets()

    Dim wksCollection As New Collection
    
    wksCollection.Add ThisWorkbook.Worksheets("VitoshAcademy")
    wksCollection.Add ThisWorkbook.Worksheets("Academy")
    wksCollection.Add ThisWorkbook.Worksheets("Vitosh")

    Dim wks As Worksheet
    Dim newWks As Worksheet

    For Each wks In wksCollection
        Dim newName As String
        newName = wks.Name & "_w"
    
        If WorksheetNameIsPresent(newName) Then
            Application.DisplayAlerts = False
            Worksheets(newName).Delete
            Application.DisplayAlerts = True
        End If
    
        wks.Copy after:=Worksheets(Worksheets.Count)
        Set newWks = Worksheets.Item(Worksheets.Count)
    
        With newWks
            .Name = newName
            .Tab.Color = vbRed
        End With
    Next wks

End Sub
