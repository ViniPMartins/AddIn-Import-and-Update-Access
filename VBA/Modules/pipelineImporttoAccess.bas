Attribute VB_Name = "pipelineImporttoAccess"
Option Explicit

Public logProgress As statusUpdate

Function pipelineUpdateAllTables()

'Esta função verifica se existem tabelas a serem atualizadas, inicia um registro de log e percorre cada tabela.
'Para cada tabela, busca o caminho do arquivo, exclui todos os dados da tabela,
'e importa os dados do arquivo correspondente.”

Dim tableNames As Variant
Dim pathTable(0) As String
Dim table As String
Dim i As Integer
Dim w As Boolean
Dim msgStatus As String
Dim log As String

    tableNames = getTableNames()
    
    If UBound(tableNames) = -1 Then
        MsgBox "Nenhuma tabela encontrada para atualização", vbCritical + vbOKOnly, "STATUS ATUALIZAÇÃO"
        Call endProcess
        Exit Function
    End If
    
    Call startLogForm
        
    For i = 0 To UBound(tableNames) - 1
    
        table = tableNames(i)
    
        pathTable(0) = buscarTextoPropriedadesTabelaAccess(table)
        
        If pathTable(0) = "" Then
            GoTo continue
        End If
        
        log = "Atualizando tabela: " & table
        Call statusUpdateLoad(log)
        
        Call deleteAllDataTable(table)
        
        If Right(pathTable(0), 5) = ".xlsx" Then
            w = walkFiles(pathTable, False, table)
        Else
            w = walkFiles(pathTable, True, table)
        End If
        
continue:
    Next i
    
    If w Then
        Call endProcess
        MsgBox "Arquivos importados com sucesso!", vbInformation + vbOKOnly, "STATUS IMPORTAÇÃO"
    Else
        Call endProcess
        MsgBox "Nenhuma tabela encontrada para atualização.", vbInformation + vbOKOnly, "STATUS IMPORTAÇÃO"
    End If

End Function

Function pipelineImportFiles(path As Variant, isfolder As Boolean, isNewTable As Boolean, table As String)

'Esta função importa arquivos para uma tabela no Access, atualizando-a se for necessário.

Dim w As Boolean
Dim folderPath As String
Dim parentFolder As String
Dim log As String
    
    Call startLogForm
    
    If isNewTable Then
        Call dropTable(table)
    End If
    
    log = "Atualizando tabela: " & table
    Call statusUpdateLoad(log)
    
    w = walkFiles(path, isfolder, table)
    
    folderPath = path(0)
    
    If isfolder Or linkArk Then
        parentFolder = folderPath
    Else
        parentFolder = ExtractFilePath(folderPath)
    End If
    
    Call gravarPropriedadesTabelaAccess(table, parentFolder)
    
    If w Then
        Call endProcess
        MsgBox "Arquivos importados com sucesso!", vbInformation + vbOKOnly, "STATUS IMPORTAÇÃO"
    End If

End Function


Function walkFiles(filesPath As Variant, isfolder As Boolean, tableImport As String)

'Esta função percorre os arquivos em um determinado caminho e importa cada arquivo para uma tabela específica.
'Se o caminho fornecido for uma pasta, ele lerá todos os arquivos com extensão '.xlsx' nessa pasta.
'Caso contrário, ele importará cada arquivo individualmente.

Dim strFile As Variant
Dim strFileComplete As String

    If VarType(filesPath) = 0 Then
        MsgBox "Nenhum arquivo selecionado para tabela " & tableImport, vbInformation + vbOKOnly, "ERRO DE IMPORTAÇÃO"
        walkFiles = False
        Exit Function
    End If
    
    If isfolder Then
        'Faz a leitura do primeiro arquivo na pasta
        strFile = Dir(filesPath(0) & "*.xlsx")
        
        Do While strFile <> ""
            strFileComplete = filesPath(0) & strFile
            Call ImportFiles(strFileComplete, tableImport)
            'Faz a leitura do próximo arquivo na pasta
            strFile = Dir
        Loop
        
    Else
        For Each strFile In filesPath
            strFileComplete = strFile
            Call ImportFiles(strFileComplete, tableImport)
        Next
    
    End If
    
    walkFiles = True

End Function

Function dropTable(strTabela As String)

'Esta função exclui uma tabela do banco de dados Access.

Dim db As Object

    Set db = appAccess.CurrentDb
    
    'Excluir a tabela atual
    On Error Resume Next
    db.Execute "DROP TABLE " & strTabela
    On Error GoTo 0

End Function

Function deleteAllDataTable(strTabela As String)

'Esta função exclui todos os dados de uma tabela no Access.

Dim db As Object

    Set db = appAccess.CurrentDb
    
    'Excluir a tabela atual
    On Error Resume Next
    db.Execute "DELETE * FROM " & strTabela
    On Error GoTo 0

End Function

Function ImportFiles(strFileComplete As String, strTabel As String)

'Esta função importa arquivos Excel para tabelas no Access.
'Caso ocorra um erro, exibe uma mensagem de erro com informações sobre o arquivo e a tabela.

Dim msg As String
    
    On Error GoTo verify
    appAccess.DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, strTabel, strFileComplete, True
    Exit Function
    
verify:
    
    msg = "Não foi possível realizar a importação das informações abaixo:" & vbNewLine & vbNewLine
    msg = msg & "Arquivo: " & strFileComplete & vbNewLine
    msg = msg & "Para tabela: " & strTabel & vbNewLine & vbNewLine
    msg = msg & "Certifique-se de que o arquivo corresponde a tabela e se os arquivos originais estão no formato padrão."
    
    MsgBox msg, vbCritical + vbOKOnly, "ERRO DE ATUALIZAÇÃO"
    
End Function

Function ExtractFilePath(FilePath As String)

'Esta função recebe um caminho de arquivo como entrada e retorna o diretório do arquivo.

    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFile As Object
    Set objFile = objFSO.GetFile(FilePath)
    ExtractFilePath = objFile.parentFolder & "\"
    
End Function

Function startLogForm()

'Esta função cria uma nova instância do formulário "statusUpdate" e o exibe em modo não modal.

    Set logProgress = New statusUpdate
    logProgress.Show vbModeless

End Function

Function statusUpdateLoad(log As String)

'"Esta função atualiza o status de carregamento exibido na janela de progresso.
'Ela adiciona uma nova mensagem ao status existente e ajusta a altura da janela de progresso.”

Dim msgStatus As String
Dim h As Double

    msgStatus = logProgress.Label1.Caption
    h = logProgress.Height
    
    With logProgress
        .Label1.Caption = msgStatus & vbNewLine & log
        .Height = h + 15
    End With
    
    DoEvents

End Function
