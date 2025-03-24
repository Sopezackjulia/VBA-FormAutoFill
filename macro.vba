Private Sub InsertButton_Click()
    ' Declaração de váriaveis
    Dim code As String
    Dim namePla As String
    Dim actualDate As String
    Dim number As String
    Dim name As String
    Dim comment As String
    Dim workShift As String
    Dim supplier As String
    Dim partName As String
    Dim partNumber As String
    Dim hours As String
    Dim partCode As String
    Dim defect As String
    Dim workerCode As String
    Dim addNumber As String
    
    ' Atribui o valor de cada caixa do formulário a variável correspondente
    code = UserForm.code.Value
    actualDate = UserForm.actualDate.Value
    number = UserForm.number.Value
    name = UserForm.name.Value
    comment = UserForm.comment.Value
    workShift = UserForm.workShift.Value
    supplier = UserForm.supplier.Value
    partName = UserForm.partName.Value
    partNumber = UserForm.partNumber.Value
    hours = UserForm.hours.Value
    partCode = UserForm.partCode.Value
    defect = UserForm.defect.Value
    workerCode = UserForm.workerCode.Value
    addNumber = UserForm.addNumber.Value
    
    ' Verificação de condições
    If code = "" Then
        MsgBox "O código não pode estar em branco!", vbExclamation, "Aviso"
    ElseIf actualDate = "" Then
        MsgBox "A data não pode estar em branco!", vbExclamation, "Aviso"
    ElseIf name = "" Then
        MsgBox "O nome do colaborador não pode estar em branco!", vbExclamation, "Aviso"
    ElseIf workShift = "" Then
        MsgBox "O turno não pode estar em branco!", vbExclamation, "Aviso"
    ElseIf hours = "" Then
        MsgBox "O tempo de mão de obra não pode estar em branco!", vbExclamation, "Aviso"
    ElseIf Len(number) < 6 Then
        MsgBox "O número do carro deve ter 6 dígitos!"
    Else
    ' Cria uma nova linha em branco
    Sheets("query").Rows(7).Insert , xlFormatFromRightOrBelow
    ' Atribui o valor de cada variável a célula correspondente na planilha
    Sheets("query").Range("A7").Value = code
    Sheets("query").Range("B7").Value = Date
    Sheets("query").Range("K7").Value = number
    Sheets("query").Range("C7").Value = name
    Sheets("query").Range("O7").Value = comment
    Sheets("query").Range("E7").Value = workShift
    Sheets("query").Range("H7").Value = supplier
    Sheets("query").Range("G7").Value = partName
    Sheets("query").Range("F7").Value = partNumber
    Sheets("query").Range("L7").Value = hours
    Sheets("query").Range("I7").Value = partCode
    Sheets("query").Range("J7").Value = defect
    Sheets("query").Range("D7").Value = workerCode
    Sheets("query").Range("P7").Value = addNumber
    ActiveWorkbook.Save
    ' Armazena o nome do workbook ativo na variável namePla
    namePla = ActiveWorkbook.name
    Sheets("Controle").Activate
    Range("T3").Value = code
    Sheets(Array("Controle")).Copy
    ' Salva o workbook ativo no SharePoint com o número de protocolo e o ano atual no nome
    ActiveWorkbook.SaveAs Filename:= _
        "https://company.sharepoint.com/sites/pasta/Protocolo-" & code & "-" & Year(Now) & ".xlsm", _
        FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False

    ' Ativa a planilha "controle", copia todas as células e cola só os valores
    Sheets(1).Activate
    Cells.Copy
    Cells.PasteSpecial xlPasteValues

    ' Salva o workbook outra vez
    ActiveWorkbook.Save

    ' Reativa o workbook original
    Workbooks(nomePla).Activate

    ' Limpa o conteúdo das células na planilha
    Sheets("Controle").Range("T3").Value = ""
    Sheets(1).Activate


    Sheets("inserir dados").Activate
    Unload UserForm
    ' Exemplo de email
    Dim OutApp As Object
    Dim OutMail As Object
    Dim text As String
    Dim title As String
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

            ' Preenche o texto, título e arquivos do email e abre o outlook para o envio
            text = "Segue em anexo os detalhes do relatório de hoje!"
            title = "Relatório"
            SharePoint = "https://company.sharepoint.com/sites/pasta/Protocolo-" & code & "-" & Year(Now) & ".xlsm"
        With OutMail
            .To = "Leader <Leader@company.com>; Supervisor <Supervisor@company.com>; Team member <Teammember@company.com>;+"
            .CC = ""
            .BCC = ""
            .Subject = title
            .Body = text
            .Attachments.Add SharePoint
            .Display
        End With
        
        ' Em caso de erro não preenche o email
        On Error GoTo fim:

        Set OutMail = Nothing
        Set OutApp = Nothing

fim:
    End If
    Exit Sub
End Sub
Private Sub nome_Change()
Dim Base As Worksheet
    Dim foundCell As Range
    ' Atribui a planilha 'base' na variável base
    Set Base = ThisWorkbook.Sheets("Base")

    ' Procura o valor selecionado no ComboBox nome na coluna A
    Set foundCell = Base.Columns("A").Find(What:=name.Value, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        ' Preenche os textbox com a descrição certa
        workerCode.Value = foundCell.Offset(0, 1).Value
    Else
     ' Caso o valor não seja encontrado deixa vazio
        gmin.Value = ""
    End If
End Sub

Private Sub partnumber_Change()
    Dim Base As Worksheet
    Dim foundCell As Range

    Set Base = ThisWorkbook.Sheets("Base")

    Set foundCell = Base.Columns("B").Find(What:=partNumber.Value, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Não exibe os alguns campos para o funcionário que está preenchendo, mas são preenchidos e salvos automaticamente de acordo com a seleção de partNumber
    partName.Visible = False
    supplier.Visible = False
    partCode.Visible = False
    
    If Not foundCell Is Nothing Then
        partName.Value = foundCell.Offset(0, 1).Value
        partCode.Value = foundCell.Offset(0, 2).Value
        supplier.Value = foundCell.Offset(0, 3).Value
    Else
        partName.Value = ""
        supplier.Value = ""
        partCode.Value = ""
    End If
End Sub
Private Sub UserForm_Initialize()
    ' Pega o valor da célula A7
    cellValue = Sheets("query").Range("A7").Value
    
    ' Atribui o valor do código como o valor da célula A7 + 1
    code.Value = cellValue + 1
    
    ' Adiciona itens ao combo box "workshift"
    Me.workShift.AddItem "1°"
    Me.workShift.AddItem "2°"

    ' Verifica se o horário é menor que 15h, se for, o turno é 1
    If Time() < "15:00" Then
    Me.workShift = "1°"
    Else
    ' Se não for, o turno é 2
    Me.workShift = "2°"
    End If
    
    ' Adiciona itens ao combo box "hours"
    Me.hours.AddItem "01:00:00"
    Me.hours.AddItem "02:00:00"
    Me.hours.AddItem "03:00:00"
    Me.hours.AddItem "04:00:00"
    Me.hours.AddItem "05:00:00"
    
    ' Preenche a variável actualDate com a data atual
    actualDate.Value = Date
    
    Dim Base As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set Base = ThisWorkbook.Sheets("Base")
    ' Pega a última linha da coluna A
    lastRow = Base.Cells(Base.Rows.Count, "A").End(xlUp).Row
    ' Faz um loop que percorre todas as linhas da coluna A
    For i = 2 To lastRow
        ' Adiciona itens ao combo box "partNumber" de acordo com os itens na célula A
        partNumber.AddItem Base.Cells(i, "A").Value
    Next i
    lastRow = Base.Cells(Base.Rows.Count, "E").End(xlUp).Row
    For i = 2 To lastRow
        name.AddItem Base.Cells(i, "E").Value
    Next i
End Sub
' Criação de função para verificar se o que está seno inserido é um número
Private Function verifyNumber(l As IReturnInteger)
    Select Case l
        Case Asc("0") To Asc("9")
            verifyNumber = l
        Case Else
            verifyNumber = 0
            MsgBox "Favor inserir apenas os 6 últimos números do carro!", vbExclamation, "CAMPO TIPO NÚMERO"
    End Select
End Function
Private Function verifyNumberPN(l As IReturnInteger)
    Select Case l
        Case Asc("0") To Asc("9")
            verifyNumberPN = l
        Case Else
            verifyNumberPN = 0
            MsgBox "Favor inserir apenas números!", vbExclamation, "CAMPO TIPO NÚMERO"
    End Select
End Function
Private Sub vin_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Chama a função de validação de números
    KeyAscii = verifyNumber(KeyAscii)
End Sub
Private Sub partnumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = verifyNumberPN(KeyAscii)
End Sub
Private Sub addNumber_Click()
    ' Exibe o campo de números adicionais se for clicado no botão de adicionar número
    addNumber.Visible = True
    addNumberLabel.Visible = True
End Sub
Private Sub cancelButton_Click()
    ' Fecha o formulário
    Unload UserForm
    Exit Sub
End Sub
