Public areatotal As Double


Function tirar_parenteses(caract)

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
    tirar_parenteses = temp
End Function

'este módulo é focado em consultar os cabos na tabela "Tabela-Cabos" a comparação contextual será usada para se achar os cabos

Private Sub calcular_area(ref As String)
    Dim c As Range
    Dim area(5) As Double
    Dim separado() As String                                                                              'separa em 2x(3x75) e 1x(2x50)
    Dim qtdcabos(5) As Integer

    
    Sheets("Tabela-Cabo").Select
    separado = Split(ref, ")+")
    
    For i = 0 To UBound(separado)
        qtdcabos(i) = Left(separado(i), 1)
        separado(i) = Mid(separado(i), 3, Len(separado(i)))
    Next

    With Sheets("Tabela-Cabo").Range("A:XFD")
        areatotal = 0
        For j = 0 To UBound(separado)
            Set c = .Find(tirar_parenteses(separado(j)), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
            area(j) = ActiveWorkbook.Worksheets("Tabela-Cabo").Range("D" & c.Row).Value * ActiveWorkbook.Worksheets("Tabela-Cabo").Range("D" & c.Row).Value
            area(j) = area(j) * qtdcabos(j)
            areatotal = areatotal + area(j)
        Next
    End With
    
End Sub
