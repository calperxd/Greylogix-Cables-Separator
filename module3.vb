Public larguratotal As Double


Function tirar_parenteses2(caract)

    codiA = "()sh"
    'Letras correspondentes para substituição
    codiB = "    "
    'Armazena em temp a string recebida
    temp = caract
    'Loop que percorerá a palavra letra a letra
    For i = 1 To Len(temp)
        'InStr buscará se a letra pertence ao grupo com acentos e se existir retornará a posição dela
        p = InStr(codiA, Mid(temp, i, 1))

        'Substitui a letra de indice i em codiA pela sua correspondente em codiB
        If p > 0 Then Mid(temp, i, 1) = Mid(codiB, p, 1)
    Next
    'Retorna o texto o caractere especial
    temp = RTrim(temp)
    temp = LTrim(temp)
    tirar_parenteses2 = temp
End Function

Private Sub procurar_largura()
    Sheets(Sheets.Count).Cells(2, 5).Select
    Dim c As Range
    Dim largura(5) As Double
    Dim separado() As String                                                                              'separa em 2x(3x75) e 1x(2x50)
    Dim qtdcabos(5) As Integer
    Max = (ActiveSheet.Cells.SpecialCells(xlLastCell).Row - 2)
    
    
    For h = 1 To Max
        ref = Split(Sheets(Sheets.Count).Cells(h + 1, 1).Value, "_")
        separado = Split(ref(2), ")+")
        
        For i = 0 To UBound(separado)
            qtdcabos(i) = Left(separado(i), 1)
            separado(i) = Mid(separado(i), 3, Len(separado(i)))
        Next
    
        With Sheets("Tabela-Cabo").Range("A:XFD")
            larguratotal = 0
            For j = 0 To UBound(separado)
                Set c = .Find(tirar_parenteses2(separado(j)), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
                largura(j) = c.Cells.Offset(0, 1)
                largura(j) = largura(j) * qtdcabos(j)
                larguratotal = larguratotal + largura(j)
                Sheets(Sheets.Count).Cells(h + 1, 5) = larguratotal
            Next j
        End With
    Next h
    Sheets(Sheets.Count).Cells((Max + 2), 5) = Application.WorksheetFunction.Sum(Worksheets(Sheets.Count).Range("E2:E" & Max + 1))
End Sub
