VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} getExternalData 
   Caption         =   "Atualizar Dados Access"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10950
   OleObjectBlob   =   "getExternalData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "getExternalData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub ComandsBottons_Click()

End Sub

Private Sub Label7_Click()

End Sub

'Inicialização do UserForm
'----------------------------------------------------------------------------------
Private Sub UserForm_Initialize()

End Sub

'Comandos dos botões
'----------------------------------------------------------------------------------
Private Sub buttonOK_Click()

If Me.optionUpdateAll.Value Then
    
    updateAll = True
    Unload Me
    Exit Sub

End If

'verificações de preenchimento dos dados
If Me.optionFile.Value = True And Me.FileText.Value = "0 arquivos selecionados" Then
    
    MsgBox "É necessário escolher pelo menos um arquivo", vbCritical + vbOKOnly, "ERRO DE PREENCHIMENTO"
    Me.FileText.SetFocus
    Exit Sub
    
End If

If Me.optionFolder.Value = True And (Me.folderText.Value = "" Or IsNumeric(Me.folderText.Value)) Then

    MsgBox "Insira um endereço de pasta válido", vbCritical + vbOKOnly, "ERRO DE PREENCHIMENTO"
    Me.folderText.SetFocus
    Exit Sub

End If

If Me.optionNewTable.Value And Me.tableName.Value = "" Then

    MsgBox "Insira um nome de tabela válido", vbCritical + vbOKOnly, "ERRO DE PREENCHIMENTO"
    Me.tableName.SetFocus
    Exit Sub

End If

If Me.optionCurrentTable.Value And Me.boxCurrentTable.Value = "" Then

    MsgBox "Selecione uma tabela válida", vbCritical + vbOKOnly, "ERRO DE PREENCHIMENTO"
    Me.boxCurrentTable.SetFocus
    Exit Sub

End If

'atribuição das variáveis

isfolder = Me.optionFolder.Value
isNewTable = Me.optionNewTable.Value
linkArk = Me.CheckBoxArk.Value

If Me.optionNewTable.Value Then

    strTableName = Me.tableName.Value

ElseIf Me.optionCurrentTable.Value Then

    strTableName = Me.boxCurrentTable.Value

End If

updateAll = False
Unload Me
'Application.Visible = True

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then Cancel = True
End Sub

Private Sub buttonCancel_Click()

Call endProcess

End Sub

Private Sub newConnectionbutton_Click()

Call newConnection
Call ConnectionDB

End Sub

Private Sub FolderBrowser_Click()

path = getImportFiles(True)

If IsEmpty(path) Then
    Exit Sub
Else
    Me.folderText.Value = path(0)
End If

End Sub


Private Sub FileBrowser_Click()

Dim msg As String

path = getImportFiles(False)

If IsEmpty(path) Then
    Exit Sub
Else
    If UBound(path) = 0 Then
        Me.CheckBoxArk.Visible = True
        msg = UBound(path) + 1 & " arquivo selecionado"
        Me.FileText.Value = msg
    Else
        Me.CheckBoxArk.Visible = False
        Me.CheckBoxArk.Value = False
        msg = UBound(path) + 1 & " arquivos selecionados"
        Me.FileText.Value = msg
    End If

End If

End Sub

'Procedimentos das opções selecionadas
'---------------------------------------------------------------------------------------------------
Private Sub optionUpdateAll_Click()

Me.SourceImport.Enabled = False

Me.optionFolder.Enabled = False
Me.folderText.Enabled = False
Me.FolderBrowser.Enabled = False

Me.optionFile.Enabled = False
Me.FileBrowser.Enabled = False

Me.TypeStoreData.Enabled = False
Me.optionNewTable.Enabled = False
Me.tableName.Enabled = False

Me.optionCurrentTable.Enabled = False
Me.boxCurrentTable.Enabled = False

Me.Alert.Visible = True
Me.msgAlert.Visible = True

End Sub

Private Sub OptionInsertDatabase_Click()

Me.SourceImport.Enabled = True

Me.optionFolder.Enabled = True
Me.folderText.Enabled = True
Me.FolderBrowser.Enabled = True

Me.optionFile.Enabled = True
Me.FileBrowser.Enabled = True

Me.TypeStoreData.Enabled = True
Me.optionNewTable.Enabled = True
Me.tableName.Enabled = True

Me.optionCurrentTable.Enabled = True
Me.boxCurrentTable.Enabled = True

Me.Alert.Visible = False
Me.msgAlert.Visible = False


End Sub

Private Sub optionFolder_Click()

Me.FileBrowser.Visible = False
Me.FileText.Visible = False
Me.CheckBoxArk.Visible = False
Me.CheckBoxArk.Value = False
Me.FileText.Value = "0 arquivos selecionados"

Me.FolderBrowser.Visible = True
Me.folderText.Visible = True

End Sub

Private Sub optionFile_Click()

Me.FileBrowser.Visible = True
Me.FileText.Visible = True

Me.FolderBrowser.Visible = False
Me.folderText.Visible = False
Me.folderText.Value = ""

End Sub

Private Sub optionNewTable_Click()

Me.boxCurrentTable.Visible = False
Me.tableName.Visible = True

End Sub

Private Sub optionCurrentTable_Click()

Me.boxCurrentTable.Visible = True
Me.tableName.Visible = False

End Sub
