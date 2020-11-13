Dim row_clicked As Integer 'segura o valor da linha que foi clickada
Dim column_clicked As Integer 'segura o valor da coluna que foi clickada


'Percorre o relatório rascunho preenchendo com os valores
'chamar função para preencher valor
Private Sub imputar_dados()
    Dim c As Range
    Dim campos(10) As String:
    campos(0) = "Nível"
    campos(1) = "Leito"
    campos(2) = "Área útil mm²"
    campos(3) = "Área ocupada mm²"
    campos(4) = "% Ocupação"
    campos(5) = "% Critério"
    campos(6) = "Peso Leito kg/m"
    campos(7) = "Comprimento m"
    campos(8) = "Peso cabos kg/m"
    campos(9) = "Peso Total kg/m"
    campos(10) = "Camadas"
    Sheets(Sheets.Count).Range("A1").Activate
    Selection.End(xlDown).Activate
    temp = ActiveCell.Row
    Dim cola As Integer: cola = 10
    Sheets("Gates-Resumo").Range("A10:A900").EntireRow.Delete
    Sheets("Rascunho").Range("G11") = Sheets("Geral-Gates").Range("R1")
    
    For j = 0 To (temp - 2)
        For i = 0 To UBound(campos)
            Dim name As String: name = campos(i)
            ' Get search range
            Dim rgSearch As Range
            Set rgSearch = Sheets("Rascunho").Range("A:G")
    
            Dim cell As Range
            Set cell = rgSearch.Find(name, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
            ' If not found then exit
            If cell Is Nothing Then
                Debug.Print "Not found"
                Exit Sub
            End If
            
    
            Select Case i
                Case 0
                    'nível
                    cell.Cells.Offset(1, 0) = Sheets(Sheets.Count).Cells(j + 2, 2)
                Case 1
                    'leito
                    cell.Cells.Offset(1, 0) = Sheets(Sheets.Count).Cells(j + 2, 3)
                Case 2
                    'area útil
                    cell.Cells.Offset(1, 0) = Sheets(Sheets.Count).Cells((j + 2), 8).Value2
                Case 3
                    'area ocupada
                    cell.Cells.Offset(1, 0) = Sheets(Sheets.Count).Cells(j + 2, 8) * (Sheets(Sheets.Count).Cells(j + 2, 9) / 100)
                Case 4
                    'ocupacao
                    cell.Cells.Offset(1, 0) = Sheets(Sheets.Count).Cells(j + 2, 9)
                Case 5
                    'critério
                    If Sheets("Rascunho").Range("E11") > Sheets("Geral-Gates").Range("R1") Then
                        Sheets("Rascunho").Range("E11").Interior.Color = vbRed
                    Else
                        Sheets("Rascunho").Range("E11").Interior.Color = vbWhite
                    End If
                Case 6
                    'peso leito kg/m
                    Dim c2 As Range
                    temp23 = Split(Sheets(Sheets.Count).Cells((j + 2), 3).Value, ".")
                    Set c2 = Sheets("Tabela-Gate").Range("A:F").Find(temp23(0))
                    cell.Cells.Offset(1, 0) = c2.Cells.Offset(0, 5)
                Case 7
                    'Comprimento em m
                    cell.Cells.Offset(1, 0) = Sheets(Sheets.Count).Cells(j + 2, 5)
                Case 8
                    'Peso cabos kg/m
                    Dim c3 As Range
                    Set c3 = Sheets("Geral-Gates").Range("A:I").Find(Sheets(Sheets.Count).Cells(j + 2, 2).Value)
                    linhap = c3.Row
                    colunap = c3.Column
                    Application.Run "Módulo1.criar_abas"
                    If aba = True Then
                    Sheets(Sheets.Count).Select
                        Sheets(Sheets.Count).Cells(1, 4).Select
                        cell.Cells.Offset(1, 0) = (Selection.End(xlDown).Value / CDbl((Sheets("Rascunho").Range("A15"))))
                        Application.DisplayAlerts = False
                        Sheets(Sheets.Count).Delete
                    End If
                Case 9
                    'Peso Total kg/m
                    Sheets("Rascunho").Range("D15").Value = Sheets("Rascunho").Range("B15").Value + Sheets("Rascunho").Range("C15").Value
                Case 10
                    'Camadas
                    Dim c4 As Range
                    Set c4 = Sheets("Geral-Gates").Range("A:I").Find(Sheets(Sheets.Count).Cells(j + 2, 2).Value)
                    linhap = c4.Row
                    colunap = c4.Column
                    Application.Run "Módulo1.criar_abas"
                    If aba = True Then
                        Sheets(Sheets.Count).Select
                        Sheets(Sheets.Count).Cells(1, 5).Select
                        temp = Split(Sheets(Sheets.Count - 1).Cells((j + 2), 3).Value, ".")
                        cell.Cells.Offset(1, 0) = Selection.End(xlDown).Value / temp(0)
                        Application.DisplayAlerts = False
                        Sheets(Sheets.Count).Delete
                        If Sheets("Rascunho").Range("G15") > 1 Then
                            Sheets("Rascunho").Range("G15").Interior.Color = vbYellow
                        Else
                            Sheets("Rascunho").Range("G15").Interior.Color = vbWhite
                        End If
                    End If
            End Select
        Next i
        Sheets("Rascunho").Range("A10:G15").Copy _
            Destination:=Sheets("Gates-Resumo").Cells(cola, 1)
        cola = cola + 6
    Next j
    Application.DisplayAlerts = False
    Sheets(Sheets.Count).Delete
    Sheets("Rascunho").Visible = False
    ActiveWindow.DisplayGridlines = False
    Sheets("Rascunho").Range("G15").Interior.Color = vbWhite
    Sheets("Rascunho").Range("E11").Interior.Color = vbWhite
    

End Sub

Private Sub SortDataWithHeader()
 Sheets(Sheets.Count).Range("A2").Select
 'Applying sort.
 With Sheets("Gates-Resumo").Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("K2:K" & Selection.End(xlDown).Row), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SetRange Range("K1:K" & Selection.End(xlDown).Row)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
 End With
End Sub


Private Sub filtrar_dados(gate As String)
    Dim rng As Range
    Dim ws As Worksheet
    
    Sheets("Geral-Gates").Range("A1").AutoFilter Field:=1, Criteria1:=gate
    If Sheets("Geral-Gates").AutoFilterMode = False Then
        Exit Sub
    End If
    Sheets("Gates-Resumo").Range("J:R").EntireColumn.Delete
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
    Sheets("Geral-Gates").AutoFilter.Range.Copy _
        Destination:=Sheets(Sheets.Count).Range("A1")
    Sheets("Geral-Gates").AutoFilterMode = False
End Sub

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
        Dim c As Range
        Dim firstAddress As String
        row_clicked = Target.Range.Application.ActiveCell.Row                        'segura o valor da linha que foi clickada
        column_clicked = Target.Range.Application.ActiveCell.Column                  'segura o valor da coluna que foi clickada
    Select Case Target.Range.Application.ActiveCell.Column
        Case 1
            Sheets("Gates-Resumo").Select
            Sheets("Gates-Resumo").Cells(8, 5).Value = Sheets("Geral-Gates").Cells(row_clicked, column_clicked).Value
            filtrar_dados (Sheets("Gates-Resumo").Cells(8, 5).Value)
            SortDataWithHeader
            imputar_dados
            Sheets("Gates-Resumo").Select
            Sheets("Gates-Resumo").Range("H1").Select
        Case 2
            linhap = Target.Application.ActiveCell.Row
            colunap = Target.Application.ActiveCell.Column
            Application.Run "Módulo1.criar_abas"
    End Select
End Sub

