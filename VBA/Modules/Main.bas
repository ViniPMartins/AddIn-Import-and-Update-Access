Attribute VB_Name = "Main"
Option Explicit

Sub ImportDatatoAccessPtBr()

'Esta fun��o verifica se h� uma conex�o com um banco de dados do Access e, em seguida,
'importa dados para o Access a partir de arquivos ou atualiza todas as tabelas.
'Caso n�o haja conex�o, oferece a op��o de iniciar uma nova conex�o.

Dim conn As Boolean
Dim time As Date
Dim interval As Date
Dim strOutput As String
Dim pathdatabase As String
Dim msg As String
    
verificar_conex�o:
    On Error Resume Next
    conn = ConnectionDB()
    
    If Not conn Then
    
        msg = "N�o h� nenhuma base de dados do Access conectada." & vbNewLine & vbNewLine
        msg = msg & "Deseja iniciar uma nova conex�o com um arquivo Access?"
    
        If MsgBox(msg, vbInformation + vbYesNo, "STATUS CONEX�O") = vbYes Then
            
            On Error Resume Next
            Call newConnection
            GoTo verificar_conex�o
            
        End If
        
        Exit Sub
    End If
    
    Call getResponseForm
    
    If updateAll Then
        If IsEmpty(isfolder) Then
            Exit Sub
        End If
        
        Call pipelineUpdateAllTables
        
    Else
        If IsEmpty(path) Then
            Exit Sub
        End If
        
        Call pipelineImportFiles(path, isfolder, isNewTable, strTableName)
        
    End If

End Sub

Function endProcess()

'Esta fun��o encerra o processo atual, desalocando recursos e liberando mem�ria.
'Ela descarrega um formul�rio chamado 'getExternalData', desconecta do banco de dados
'e descarrega outro formul�rio chamado 'logProgress'.

    On Error Resume Next
    path = Empty
    isfolder = Empty
    isNewTable = Empty
    strTableName = Empty
    
    Unload getExternalData
    Call DisconnectDB
    Unload logProgress
    
End Function

Function newConnection()

'Esta fun��o cria uma conex�o com um banco de dados do Access, permitindo especificar o caminho do arquivo de
'banco de dados atrav�s de uma caixa de di�logo. O caminho do arquivo � armazenado em uma propriedade
'personalizada chamada "LinkToAccess" no documento ativo.

Dim pathdatabase As String

On Error Resume Next
pathdatabase = getImportFiles(False, False, "Access Database")(0)

If pathdatabase = "" Then
    Exit Function
End If

ActiveWorkbook.CustomDocumentProperties("LinkToAccess").Delete

With ActiveWorkbook.CustomDocumentProperties
        .Add Name:="LinkToAccess", _
             LinkToContent:=False, _
             Type:=msoPropertyTypeString, _
             Value:=pathdatabase
End With

End Function
