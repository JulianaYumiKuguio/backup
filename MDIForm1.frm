VERSION 5.00
Begin VB.MDIForm frm_Principal 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Gerenciamento de Transa��es (Cr�dito)"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   10080
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Menu frmCadastro 
      Caption         =   "Transa��es"
   End
   Begin VB.Menu frmConsulta 
      Caption         =   "Consulta"
   End
   Begin VB.Menu frmRelatorio 
      Caption         =   "Relat�rio"
   End
End
Attribute VB_Name = "frm_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub frmCadastro_Click()
    frm_CadastroTransacoes.Show
End Sub

Private Sub MDIForm_Activate()
    Call AbreConexaoBD
End Sub

Sub AbreConexaoBD()

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

' Vari�vel de conex�o com o BD
'Dim cn As New ADODB.Connection
' Vari�vel de acesso a Tabela do BD
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
