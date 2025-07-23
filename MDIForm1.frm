VERSION 5.00
Begin VB.MDIForm frm_Principal 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Gerenciamento de Transações (Crédito)"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   10080
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Menu frmCadastro 
      Caption         =   "Transações"
   End
   Begin VB.Menu frmConsulta 
      Caption         =   "Consulta"
   End
   Begin VB.Menu frmRelatorio 
      Caption         =   "Relatório"
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

Private Sub frmConsulta_Click()
    frm_Consulta.Show
End Sub

Private Sub frmRelatorio_Click()
    frm_Relatorio.Show
End Sub

Private Sub MDIForm_Activate()
On Error GoTo TratamentoDeErro

 ' Call AbreConexaoBD
    
    
    
Exit Sub
TratamentoDeErro:
    ' Monta a mensagem de log com detalhes do erro
    Dim strErroDetails As String
    strErroDetails = "Erro na rotina MinhaRotinaQuePodeGerarErro - " & _
                     "Número: " & Err.Number & " | " & _
                     "Descrição: " & Err.Description & " | " & _
                     "Fonte: " & Err.Source & " | " & _
                     "ÚltimaDLL: " & Err.HelpFile & " | " & _
                     "Contexto: Linha do erro/Estado da aplicação" ' Adicione contexto se possível

    ' Chama a rotina de log do módulo1
    Call EscreverLogErro(strErroDetails)

    ' Opcional: Avisar o usuário de forma amigável (sem mostrar detalhes técnicos)
    MsgBox "Ocorreu um erro inesperado. O problema foi registrado e será investigado.", vbCritical, "Erro"


End Sub

Sub AbreConexaoBD()

f = FreeFile
Open App.Path & "\Config.ini" For Input As f    'abre o arquivo texto

On Error Resume Next    'se a tabela não existir escapa da mensagem de erro

On Error GoTo trata_erro   'ativa tratamento de erros

Line Input #f, linha 'lê uma linha do arquivo texto
Do While Not EOF(f)
    ''Extrai a informação do arquivo .ini
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

' Variável de conexão com o BD
'Dim cn As New ADODB.Connection
' Variável de acesso a Tabela do BD
 Dim rs As New ADODB.Recordset


    cn.Provider = "SQLOLEDB"    ' Provedor de acesso ao SQL Server
    cn.Properties("Data Source").Value = xServer
    cn.Properties("Initial Catalog").Value = xDataBase
    cn.Properties("User ID").Value = xUserName
    cn.Properties("Password").Value = xPassword
    If rs.State = adStateOpen Then rs.Close
    cn.Open  ' Abrindo a conexão

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = cn

Exit Sub

trata_erro:
MsgBox Err.Description
End Sub

Private Sub teste_Click()
    Form1.Show
End Sub
