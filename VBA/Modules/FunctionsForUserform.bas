Attribute VB_Name = "FunctionsForUserform"
Option Explicit

Public path As Variant
Public isfolder As Boolean
Public isNewTable As Boolean
Public strTableName As String
Public updateAll As Boolean
Public linkArk As Boolean

Function getResponseForm()

'Esta função carrega uma janela de formulário chamada 'getExternalData' e preenche um combobox com nomes de tabelas.
'A janela é exibida na posição 2 (centralizada).”

Dim tableNames As Variant
Dim table As Variant

    tableNames = getTableNames()
    
    Load getExternalData
    
    For Each table In tableNames
        getExternalData.boxCurrentTable.AddItem table
    Next
    
    getExternalData.StartUpPosition = 2
    getExternalData.Show

End Function

Function getTableNames()

'Esta função retorna os nomes das tabelas de um banco de dados Access.

Dim db As Object
Dim tbl As Object
Dim i As Integer
Dim n As Integer
Dim tableNames As Variant
Dim strTables As String
    
    Set db = appAccess.CurrentDb
    
    strTables = ""
    i = 0
    
    For Each tbl In db.tabledefs
        If Left(tbl.Name, 4) <> "MSys" Then
            strTables = strTables & tbl.Name & ";"
            i = i + 1
        End If
    Next tbl
    
    tableNames = Split(strTables, ";")
    getTableNames = tableNames
    
End Function

Function getImportFiles(isfolder As Boolean, Optional multiSelect As Boolean = True, Optional filesType As String = "Excel Files")

'Esta função obtém os arquivos de importação, permitindo selecionar uma pasta ou arquivos individuais.

Dim fDialog As Office.FileDialog
Dim varFile As Variant
Dim varFolder As Variant
Dim source() As String
Dim nFiles As Integer
Dim totalallFiles As Integer
Dim tlt As String
Dim tp As String

    ' Set up the File Dialog.
    If isfolder Then
        Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
        
        With fDialog
            .AllowMultiSelect = False
            tlt = "Selecione a pasta que deseja importar"
            .Title = tlt
            
            If .Show = True Then
                ReDim source(1)
                For Each varFolder In .SelectedItems
                    source(0) = varFolder & "\"
                Next
                getImportFiles = source
                
            Else
            
                Exit Function
                
            End If
        End With
                
        
    Else
        Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
        
        Select Case filesType
        
        Case "Access Database"
            tp = "*.accdb"
            
        Case "Excel Files"
            tp = "*.xls; *.xlsx"
            
        Case Else
            filesType = "All Files"
            tp = "*.*"
            
        End Select
    
        With fDialog
            .AllowMultiSelect = multiSelect
            tlt = "Selecione o arquivo que deseja importar"
            .Title = tlt
            .Filters.Clear
            .Filters.Add filesType, tp
            
            If .Show = True Then
                totalallFiles = .SelectedItems.Count - 1
                ReDim source(totalallFiles)
                nFiles = 0
                
                For Each varFile In .SelectedItems
                    source(nFiles) = varFile
                    'Debug.Print allFiles(nFiles)
                    nFiles = nFiles + 1
                Next
                getImportFiles = source
                
            Else
            
                Exit Function
                
            End If
        End With

    End If

End Function

