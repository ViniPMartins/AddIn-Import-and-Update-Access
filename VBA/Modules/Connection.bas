Attribute VB_Name = "Connection"
Public appAccess As Object

Function ConnectionDB()

'Esta função cria uma conexão com um banco de dados do Access usando o caminho especificado em uma
'propriedade personalizada com o nome "LinkToAccess".

    On Error Resume Next
    With statusUpdate
        .Caption = "Conexão"
        .Label1 = "Conectando a base de dados Access..."
        .Show vbModeless
    End With
    
    pathdatabase = ActiveWorkbook.CustomDocumentProperties("LinkToAccess").Value
    
    Set appAccess = CreateObject("Access.Application")
    
    On Error GoTo alter
        appAccess.OpenCurrentDatabase pathdatabase
        ConnectionDB = True
        Unload statusUpdate
        Exit Function
    
alter:
    ConnectionDB = False
    Unload statusUpdate

End Function

Function DisconnectDB()

'Esta função desconecta o banco de dados atualmente aberto no aplicativo Access,
'encerrando a aplicação e liberando os recursos associados.

    On Error Resume Next
    appAccess.CloseCurrentDatabase
    appAccess.Quit
    Set appAccess = Nothing

End Function

