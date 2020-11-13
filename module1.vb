'Variáveis globais

Public nome As String                          'recebe o valor com caracteres especiais
Dim nome1 As String                         'recebe o valor sem caracteres especiais
Public contaLinhas As Long
Public linhap As Integer                    'pega a linha do hiperlink clickado
Public colunap As Integer                   'pega a coluna do hiperlink clickado
Public dimensao As String                   'esta variável recebe a dimensão do cabo depois do terceiro underline ou seja 1x(3x23)+3(2x35) etc
Public aba As Boolean                       'false não criar aba true para criar aba
Public taxaocup As Double                   'variavel para guardar a taxa de ocupação temporária durante a cópia de taxas
Public naoestoura As Boolean                'variável que não deixa dar estouro de pilha, quando houver apenas 1 cabo dentro do gate uma vez que a contagem de linhas é feita através de xlDown, ele não deixa estourar a pilha
'End

Function tirar_mm(caract)

    codiA = "mm"
     
    'Letras correspondentes para substituição
    codiB = " "

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
    tirar_mm = Trim(temp)

End Function


Function especial(caract)

    codiA = "/"
     
    'Letras correspondentes para substituição
    codiB = " "

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
    especial = Trim(temp)

End Function

'função para tirar a barra do começo do nome
Function tirar_barra(caract)

    codiA = "/"
     
    'Letras correspondentes para substituição
    codiB = ""

    'Armazena em temp a string recebida
    temp = caract

    'Loop que percorerá a palavra letra a letra
    For i = 1 To 1
        'InStr buscará se a letra pertence ao grupo com acentos e se existir retornará a posição dela
        p = InStr(codiA, Mid(temp, i, 1))

        'Substitui copia a string sem a barra
        If p > 0 Then
        
            temp = Mid(temp, p + 1, Len(caract))                        'Mid(temp, i, 1) = Mid(codiB, p, 1)
        
        End If
    Next

    'Retorna o texto o caractere especial
    tirar_barra = temp

End Function

'taxa de ocupação
Private Sub taxa_de_ocup()
    Dim contalinha As Integer                                                                   'conta o número máximo de linhas da nova aba
    Dim area As Double
    Cells(2, 1).Activate                                                                        'seleciona a linha 2 e coluna 1 para iniciar a varredura de linhas
    Worksheets(Sheets.Count).Activate
    ActiveSheet.Cells.SpecialCells(xlLastCell).Activate
    contalinha = ActiveCell.Row
    area = Sheets("Geral-Gates").Cells(linhap, 8).Value
    Sheets(Sheets.Count).Activate
    For i = 2 To contalinha - 1
        Cells(i, 3).Value = ((Cells(i, 2) / area) * 100)
    Next i
    Sheets(Sheets.Count).Cells(contalinha, 3).Value = Application.WorksheetFunction.Sum(Worksheets(Sheets.Count).Range("C2:C" & contalinha + 1))
End Sub



'função para procurar cabos e retornar referência
Private Sub procurar_cabos(gate As String)

    Dim primeiraocorrencia As Range
    Dim ultimaocorrencia As Range
    Dim iterador As String
    Dim c As Range
    Dim i As Integer
    
    i = 2           'começa da linha 2

    With Sheets(2).Range("A:Z")
    
        Set c = .Find(tirar_barra(gate), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
        Set primeiraocorrencia = c                                                                      'primeiro item achado
        Set c = .Find(tirar_barra(gate), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
        Set ultimaocorrencia = c                                                                        'último item
        If Not c Is Nothing Then
            Do

                Set c = .FindNext(c)
                iterador = c.Address
                ActiveWorkbook.Worksheets(2).Range("A" & c.Row).Copy                                       'copia da planilha de cabos as células
                ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count).Activate                          'vai para a última planilha criada
                Worksheets(Sheets.Count).Range("A:Z").Columns.AutoFit
                ActiveSheet.Paste Destination:=Worksheets(ActiveWorkbook.Worksheets.Count).Range("A" & i)  'cola na última planilha criada
                
                i = i + 1
                temp = StrComp(iterador, ultimaocorrencia.Address, 0)                                       'guarda o resultado para saber se está no último item achado
                If temp = 0 Then                                                                            'indica que chegou na última célula encontrada
                    Set c = Nothing
                End If
            Loop While Not c Is Nothing
        End If
    End With
End Sub

'função para procurar largura do cabo quadrado
Private Sub procurar_area_util()
    Dim contalinha As Integer                                                                   'conta o número máximo de linhas da nova aba
    Dim temp() As String                                                                          'variável temporária que recebe o valor da célula naquele momento
    Cells(2, 1).Activate                                                                        'seleciona a linha 2 e coluna 1 para iniciar a varredura de linhas
    Worksheets(Sheets.Count).Activate
    ActiveSheet.Cells.SpecialCells(xlLastCell).Activate
    contalinha = ActiveCell.Row
    Cells(2, 1).Activate                                                                        'volta para a primeira célula
    For j = 2 To contalinha
        temp = Split(Sheets(Sheets.Count).Cells(j, 1).Value, "_")
        dimensao = temp(2)
        Application.Run "Módulo2.calcular_area", dimensao
        Sheets(Sheets.Count).Activate
        Cells(j, 2).Value = areatotal
    Next j
        Sheets(Sheets.Count).Activate
        Sheets(Sheets.Count).Cells(contalinha + 1, 2).Value = Application.WorksheetFunction.Sum(Worksheets(Sheets.Count).Range("B2:B" & contalinha + 1))

End Sub

'preencher área nas Geral-Gates
Private Sub preencher_area_total_gates()
    Dim c As Range
    Dim temp() As String
    Sheets("Geral-Gates").Range("H" & 1).Value = "Área útil mm²"
    Cells(2, 8).Activate
    For i = 2 To contaLinhas
        temp = Split(Sheets("Geral-Gates").Range("C" & i).Value, ".")
        Set c = Sheets("Tabela-Gate").Range("A:A").Find(temp(0))
        If Not c Is Nothing Then
            Sheets("Geral-Gates").Range("H" & i).Value = c.Cells.Offset(0, 4)
            Set c = Nothing
        End If
    Next i
End Sub

'função para calcular o peso do Cabo
Private Sub calcular_peso_cabo()
    Dim c As Range
    Dim pesokgkm As Double                                                                      ' pega o kg/km do cabo
    Dim extensaogate As Double                                                                  ' pega a extensão do gate
    Dim separado() As String                                                                              'separa em 2x(3x75) e 1x(2x50)
    Dim qtdcabos(5) As Integer
    Dim peso(5) As Double
    Dim pesototal As Double
    contalinha = ActiveSheet.Cells.SpecialCells(xlLastCell).Row
    'resolver amanhã esse problema daqui para baixo
    'função deve passar um parâmetro aqui , então verificar a possibilidade de adaptar uma outra função que seja equivalente à essa , mas que não precise de um loop
    For j = 0 To ActiveSheet.Cells.SpecialCells(xlLastCell).Row - 2 Step 1
        If (contalinha - 2) <> j Then
            temp = Split(Cells(j + 2, 1).Value, "_")                        'variável que tem a informação dos cabos
            separado = Split(temp(2), ")+")
            For i = 0 To UBound(separado)
                qtdcabos(i) = Left(separado(i), 1)
                separado(i) = Mid(separado(i), 3, Len(separado(i)))
            Next i
            
            'este laço calcula o pesototal do cabo ao longo do gate
            For m = 0 To UBound(separado)
                Set c = Sheets("Tabela-Cabo").Range("A:XFD").Find(tirar_parenteses(separado(m)), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
                pesokgkm = Sheets("Tabela-Cabo").Range("F" & c.Row).Value
                Set c = Sheets("Geral-Gates").Range("A:XFD").Find(Sheets(Sheets.Count).Cells(1, 1).Value, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
                extensaogate = Sheets("Geral-Gates").Range("E" & c.Row).Value
                peso(m) = pesokgkm * extensaogate * qtdcabos(m)
                pesototal = pesototal + peso(m)
            Next m
            Sheets(Sheets.Count).Cells(j + 2, 4).Value = (pesototal / 1000)
            pesototal = 0
        End If
    Next j
    Sheets(Sheets.Count).Cells(contalinha, 4).Value = Application.WorksheetFunction.Sum(Worksheets(Sheets.Count).Range("D2:D" & contalinha + 1))
End Sub

'função para criar colunas e preencher com os cabos referentes as colunas
Private Sub criar_colunas()
    For i = 1 To 5 Step 1
    ActiveSheet.Cells(1, i).Activate
        Select Case i
            Case 1
                ActiveSheet.Cells(1, i).Value = nome
                procurar_cabos (nome)
            Case 2
                ActiveSheet.Cells(1, i).Value = "Área ocupada mm²"
                procurar_area_util
            Case 3
                ActiveSheet.Cells(1, i).Value = "Taxa de Ocupação %"
                taxa_de_ocup
            Case 4
                ActiveSheet.Cells(1, i).Value = "Peso do Cabo kg"
                calcular_peso_cabo
            Case 5
                ActiveSheet.Cells(1, i).Value = "Largura"
                Worksheets(Sheets.Count).Range("A:XFD").Columns.AutoFit
                Application.Run "Módulo3.procurar_largura"
            Case Else
        End Select
    Next i

End Sub

'função para calcular a taxa de ocupacao
Private Sub criar_abas()
        Dim c As Range
        
        'este bloco garante que o elemento que estamos procurando existe na planilha de cabos , caso contrario ele não cria e nem no elemento
        nome = Sheets("Geral-Gates").Cells(linhap, colunap)
        nome1 = especial(nome)
        Set c = Sheets("Cabo-Rota").Range("A:XFD").Find(Right(nome, Len(nome) - 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
        If Not c Is Nothing Then
            aba = True                                                                          'existe elemento
        Else
            aba = False                                                                         'não existe elemento
        End If
        If aba = True Then
            Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)                'cria uma nova planilha
            Sheets("Geral-Gates").Activate                                                              'seleciona a planilha Geral-Gates
            nome = ActiveSheet.Cells(linhap, colunap)                                                   'recebe a string da linha 2 até o final que é contaLinhas
            nome1 = especial(nome)                                                                      'nome1 recebe nome para retirar caracteres especiais
                ' condição para indicar se ja existe alguma planilha com o mesmo nome
                For j = 1 To Sheets.Count
                    temp = StrComp(ActiveWorkbook.Worksheets(j).name, nome1, 0)
                    If temp = 0 Then
                        j = Sheets.Count
                    End If
                Next j
                If temp <> 0 Then
                    Sheets(Sheets.Count).Activate                                                       'seleciona a última planilha
                    ActiveSheet.name = nome1                                                            'renomeia a aba
                    criar_colunas                                                                       'criar as colunas
                End If
        End If
        Sheets("Geral-Gates").Activate                                                              'voltar para a planilha Geral-Gates
End Sub

'criar hiperlinks na planilha Geral-Gates
Private Sub criar_hiperlinks()
    With Worksheets("Geral-Gates")
        .Hyperlinks.Add Anchor:=.Range("A2:B" & contaLinhas), _
        Address:="", _
        ScreenTip:=""
    End With
End Sub

'função para contabilizar número máximo de linhas da planilha Geral-Gates
Private Sub contar_linhas()
    'declarando variáveis locais
    Dim i As Integer
    'END
    'congela a tela do Excel para não ficar mostrando execução de linha a linha da macro
    'contando quantas células (colunas) estão preenchidas a partir da B1, leitura de cima para baixo.
    Sheets("Geral-Gates").Cells(1, 1).Activate                                                                'selecionando Planilha
    contaLinhas = Selection.End(xlDown).Row                                                                   'selecionando a última célula
    Sheets("Geral-Gates").Activate                                                                            'selecionando Planilha a planilha principal
End Sub

'função para remover espaços da Tabela-Cabos
Private Sub remove_espaços_tabela_cabos()
    Dim ultima As Integer                                                                            'última linha da planilha de consulta de cabos
    Dim rLocal As Range
    Sheets("Tabela-Cabo").Select
    Sheets("Tabela-Cabo").Range("C3").Select                                                        'selecionando Planilha
    ultima = Selection.End(xlDown).Row
    For i = 3 To ultima
        Set rLocal = Cells(i, 3)
        rLocal.Replace What:=" ", Replacement:=""
    Next
    Sheets("Geral-Gates").Activate                                                                    'selecionando Planilha principal
End Sub

'função para criar e deletar as planilhas
Private Sub calcular_area_sumario()

    For i = 2 To contaLinhas
        linhap = i
        colunap = 2
        criar_abas
        If aba = True Then
            Sheets(Sheets.Count).Activate
            Range("C2").Activate
            Selection.End(xlDown).Activate
            taxaocup = ActiveCell.Value
            Application.DisplayAlerts = False
            Sheets(Sheets.Count).Delete
            Sheets("Geral-Gates").Activate
            Cells(i, 9).Value = taxaocup
        End If
    Next
    Sheets("Geral-Gates").Range("I1").Value = "Taxa de ocupação %"
    Worksheets("Geral-Gates").Range("A:Z").Columns.AutoFit
End Sub

'função para remover espaços da Tabela-Cabos
Private Sub troca_pontos_por_virgula()
    Dim ultima As Integer                                                                            'última linha da planilha de consulta de cabos
    Dim rLocal As Range
    
    Sheets("Geral-Gates").Cells(2, 5).Activate                                                                   'selecionando Planilha
    ultima = Selection.End(xlDown).Row
    For i = 2 To ultima
        Set rLocal = Cells(i, 5)
        rLocal.Replace What:=".", Replacement:=","
    Next
    Sheets("Geral-Gates").Activate                                                                    'selecionando Planilha principal
End Sub

'Criação de novas abas com nomes pre-estabelecidos
Sub main()

    contar_linhas
    criar_hiperlinks
    preencher_area_total_gates
    remove_espaços_tabela_cabos
    troca_pontos_por_virgula
    calcular_area_sumario
    
End Sub

