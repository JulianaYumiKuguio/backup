Attribute VB_Name = "Module1"
Option Explicit


Global cn  As New ADODB.Connection

Public xServer, xUserName, xPassword, xDataBase, xCaminhoRPT, xCaminhoLOG As String
Public nUsuario As Integer
Public nNomeUsuario As String
Public nAdministrador As Boolean
Public WxRetorno(0 To 50)

Dim f As Long, linha As String

Sub AbreConexaoBD()

    '' Abre a conex�o Atrav�s da configura��o do arquivo config
    '' Onde � informado os par�metros (Servidor, Usu�rio, Senha, Banco de Dados, Caminho do RPT Crystal, Caminho do Log de Erros)
    '' Permitindo n�o deixar a conex�o do SQL fixo no c�digo, facilitando a manuten��o
    '' O arquivo deve ficar na mesma pasta do projeto
      
    
    f = FreeFile
    Open App.Path & "\Config.ini" For Input As f    'abre o arquivo texto
    
    On Error Resume Next    'se a tabela n�o existir escapa da mensagem de erro
    
    On Error GoTo trata_erro   'ativa tratamento de erros
    
    Line Input #f, linha 'l� uma linha do arquivo texto
    Do While Not EOF(f)
        ''Extrai a informa��o do arquivo .ini
        If linha = "[SERVER]" Then
            Line Input #f, linha
            xServer = linha
        ElseIf linha = "[USERNAME]" Then
            Line Input #f, linha
            xUserName = linha
        ElseIf linha = "[PASSWORD]" Then
            Line Input #f, linha
            xPassword = linha
        ElseIf linha = "[DATABASE]" Then
            Line Input #f, linha
            xDataBase = linha
        ElseIf linha = "[RPT]" Then
            Line Input #f, linha
            xCaminhoRPT = linha
        ElseIf linha = "[LOG]" Then
            Line Input #f, linha
            xCaminhoLOG = linha
        Else
            Line Input #f, linha
        End If
    Loop
      
    
    Close #f
    
    '' Vari�vel de conex�o com o BD
    '' Dim cn As New ADODB.Connection
    '' Vari�vel de acesso a Tabela do BD
    
     Dim rs As New ADODB.Recordset
    
        cn.Provider = "SQLOLEDB"    ' Provedor de acesso ao SQL Server
        cn.Properties("Data Source").Value = xServer
        cn.Properties("Initial Catalog").Value = xDataBase
        cn.Properties("User ID").Value = xUserName
        cn.Properties("Password").Value = xPassword
        cn.Open  ' Abrindo a conex�o
       
        Set rs = New ADODB.Recordset
        Set rs.ActiveConnection = cn
    
    Exit Sub

trata_erro:
MsgBox Err.Description
End Sub


Public Sub EscreverLogErro(ByVal strMensagem As String)

    '' Aqui criei a rotina para registrar Log de Erros em arquivo (.txt)
    '' Quando o erro acontecer inesperadamente,
    '' o sistema ir� registrar no arquivo txt (MinhaApp_LogErros.txt) no C:\Temp do computador do Usu�rio
    '' conforme configurado no arquivo Config.ini [LOG]
    
    
    Dim intFileNum As Integer
    Dim strLogEntry As String

    On Error GoTo TratarErroInterno ' Tratamento de erro para a pr�pria rotina de log

    '' Obt�m um n�mero de arquivo livre
    intFileNum = FreeFile

    '' Formata a entrada do log com data, hora e a mensagem do erro
    strLogEntry = Format(Now, "dd/mm/yyyy hh:nn:ss") & " - " & strMensagem

    '' Abre o arquivo em modo Append (adiciona ao final)
    '' Se o arquivo n�o existir, ele ser� criado
    Open xCaminhoLOG For Append As #intFileNum

    '' Escreve a entrada no arquivo
    Print #intFileNum, strLogEntry

    '' Fecha o arquivo
    Close #intFileNum

    Exit Sub

TratarErroInterno:
    '' Este � um erro dentro da pr�pria rotina de log.
    Debug.Print "Erro ao escrever no arquivo de log: " & Err.Description
    If intFileNum <> 0 Then Close #intFileNum ' Garante que o arquivo seja fechado se aberto
End Sub
