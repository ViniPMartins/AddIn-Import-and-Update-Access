Attribute VB_Name = "recordMetadataConnection"
Function gravarPropriedadesTabelaAccess(tableName As String, pathToRecord As String)

'Esta função grava o caminho de uma pasta ou arquivo na propriedade "validationText" de uma tabela do Access.

Dim db As Object
Dim tbl As Object

Set db = appAccess.CurrentDb

For Each tbl In db.tabledefs
    
    If tbl.Name = tableName Then
        tbl.ValidationText = pathToRecord
    End If
    
Next

End Function

Function buscarTextoPropriedadesTabelaAccess(tableName As String)

'Esta função busca o texto com o caminho de uma pasta ou arquivo na propriedade "validationText" de uma tabela do Access.

Dim db As Object
Dim tbl As Object

Set db = appAccess.CurrentDb

For Each tbl In db.tabledefs
    
    If tbl.Name = tableName Then
        buscarTextoPropriedadesTabelaAccess = tbl.ValidationText
        Exit Function
    End If
    
Next

End Function

