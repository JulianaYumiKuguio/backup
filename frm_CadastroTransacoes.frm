VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_CadastroTransacoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Transa��es"
   ClientHeight    =   6840
   ClientLeft      =   3750
   ClientTop       =   3195
   ClientWidth     =   7740
   Icon            =   "frm_CadastroTransacoes.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7740
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   1588
      ButtonWidth     =   1111
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Novo"
            Key             =   "Novo"
            Object.ToolTipText     =   "Novo/Limpar Dados"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar Dados"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Excluir"
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir Dados"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fechar"
            Key             =   "Fechar"
            Object.ToolTipText     =   "Fechar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      OLEDropMode     =   1
   End
   Begin VB.CommandButton cmd_Buscar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Picture         =   "frm_CadastroTransacoes.frx":1084A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Buscar Transa��es"
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txt_Data_Cadastro 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5040
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txt_CodTransacao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   6975
      Begin VB.TextBox txt_Descricao 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   3240
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         ToolTipText     =   "at� 255 Caracteres."
         Top             =   1200
         Width           =   3015
      End
      Begin VB.ComboBox cmbStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frm_CadastroTransacoes.frx":10A7D
         Left            =   3240
         List            =   "frm_CadastroTransacoes.frx":10A8A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3720
         Width           =   2775
      End
      Begin VB.TextBox txt_Valor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         ToolTipText     =   "Somente decimais positivos"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txt_NumeroCartao 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         MaxLength       =   16
         TabIndex        =   1
         ToolTipText     =   "16 d�gitos"
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "* Status :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "* Descri��o :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1680
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "* Valor (R$) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "* N�mero Cart�o :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Aten��o! Campos Obrigat�rios (*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Transa��o:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Id.Transa��o:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   9360
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_CadastroTransacoes.frx":10AB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_CadastroTransacoes.frx":119CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_CadastroTransacoes.frx":12221
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_CadastroTransacoes.frx":12FF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_CadastroTransacoes.frx":13D45
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_CadastroTransacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Buscar_Click()

On Error GoTo TratamentoDeErro
    '' T�tulo DataGrid
    frmLocalizar.lstTitulo.List(0) = "Id_Transacao"
    frmLocalizar.lstTitulo.List(1) = "Numero_Cartao"
    frmLocalizar.lstTitulo.List(2) = "Data_Transacao"
    frmLocalizar.lstTitulo.List(3) = "Status"
    frmLocalizar.lstTitulo.List(4) = "Valor_Transacao"
    frmLocalizar.lstTitulo.List(5) = "Descricao"
    
    '' T�tulo BD
    frmLocalizar.lstCampo.List(0) = "Id_Transacao"
    frmLocalizar.lstCampo.List(1) = "Numero_Cartao"
    frmLocalizar.lstCampo.List(2) = "CONVERT(VARCHAR(10), Data_Transacao, 103)"
    frmLocalizar.lstCampo.List(3) = "CASE WHEN Status = 1 THEN 'Aprovada' WHEN Status = 2 THEN 'Pendente' ELSE 'Cancelada' end"
    frmLocalizar.lstCampo.List(4) = "replace(CONVERT(NUMERIC(18, 2), Valor_Transacao ),'.',',')"
    frmLocalizar.lstCampo.List(5) = "Descricao"
    
    '' Par�metros Tabela BD
    frmLocalizar.txtTabela.Text = "tb_Transacoes"
    frmLocalizar.lstOrdem.List(0) = "Id_Transacao"
    frmLocalizar.lstRetorna.List(0) = "0"
    frmLocalizar.FMontaChave
    frmLocalizar.Caption = " Buscar Transa��es"
    frmLocalizar.Show 1
    If WxRetorno(0) <> Empty Then
        txt_CodTransacao.Text = WxRetorno(0)
        Call txt_CodTransacao_LostFocus
    End If
    
     
    
Exit Sub
TratamentoDeErro:
    '' Monta a mensagem de log com detalhes do erro
    Dim strErroDetails As String
    strErroDetails = "Erro na rotina MinhaRotinaQuePodeGerarErro - " & _
                     "N�mero: " & Err.Number & " | " & _
                     "Descri��o: " & Err.Description & " | " & _
                     "Fonte: " & Err.Source & " | " & _
                     "�ltimaDLL: " & Err.HelpFile & " | " & _
                     "Contexto: Linha do erro/Estado da aplica��o" ' Adicione contexto se poss�vel

    '' Chama a rotina de log do m�dulo1
    Call EscreverLogErro(strErroDetails)

    '' Opcional: Avisar o usu�rio de forma amig�vel (sem mostrar detalhes t�cnicos)
    MsgBox "Ocorreu um erro inesperado. O problema foi registrado e ser� investigado.", vbCritical, "Erro"

End Sub

Private Sub Form_Load()

On Error GoTo TratamentoDeErro

    '' Renomeia t�tulo da Tela e Limpa os Campos
    Me.Caption = "XYZ - Administradora de Cart�es de Cr�dito - " + Me.Caption
    Call fLimparCampos
    
    
    
Exit Sub
TratamentoDeErro:
    '' Monta a mensagem de log com detalhes do erro
    Dim strErroDetails As String
    strErroDetails = "Erro na rotina MinhaRotinaQuePodeGerarErro - " & _
                     "N�mero: " & Err.Number & " | " & _
                     "Descri��o: " & Err.Description & " | " & _
                     "Fonte: " & Err.Source & " | " & _
                     "�ltimaDLL: " & Err.HelpFile & " | " & _
                     "Contexto: Linha do erro/Estado da aplica��o" ' Adicione contexto se poss�vel

    '' Chama a rotina de log do m�dulo1
    Call EscreverLogErro(strErroDetails)

    '' Opcional: Avisar o usu�rio de forma amig�vel (sem mostrar detalhes t�cnicos)
    MsgBox "Ocorreu um erro inesperado. O problema foi registrado e ser� investigado.", vbCritical, "Erro"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

On Error GoTo TratamentoDeErro

Select Case Button.Key

    Case "Novo"
            Call fLimparCampos
     
     
    Case "Salvar"
            If Not fInconsistencias Then
               Call fGravar
            End If
       
       
    Case "Excluir"
            If Wval(txt_CodTransacao) > 0 Then
                Call fExcluir
            Else
                MsgBox "� necess�rio um Registro para Exclus�o!", vbInformation, "Aten��o."
            End If
        
        
    Case "Fechar"
            Unload Me
    
End Select

Exit Sub
TratamentoDeErro:
    '' Monta a mensagem de log com detalhes do erro
    Dim strErroDetails As String
    strErroDetails = "Erro na rotina MinhaRotinaQuePodeGerarErro - " & _
                     "N�mero: " & Err.Number & " | " & _
                     "Descri��o: " & Err.Description & " | " & _
                     "Fonte: " & Err.Source & " | " & _
                     "�ltimaDLL: " & Err.HelpFile & " | " & _
                     "Contexto: Linha do erro/Estado da aplica��o" ' Adicione contexto se poss�vel

    '' Chama a rotina de log do m�dulo1
    Call EscreverLogErro(strErroDetails)

    '' Opcional: Avisar o usu�rio de forma amig�vel (sem mostrar detalhes t�cnicos)
    MsgBox "Ocorreu um erro inesperado. O problema foi registrado e ser� investigado.", vbCritical, "Erro"


End Sub

Private Function fInconsistencias()

fInconsistencias = True


    '' Valida campos obrigat�rios

   If txt_NumeroCartao.Text = Empty Then
      MsgBox "N�mero do Cart�o � Obrigat�rio!", vbInformation, "Aten��o."
      txt_NumeroCartao.SetFocus
      Exit Function
   End If
   
   If txt_Valor.Text = Empty Then
      MsgBox "Valor � Obrigat�rio!", vbInformation, "Aten��o."
      txt_Valor.SetFocus
      Exit Function
   End If
   
   If txt_Descricao.Text = Empty Then
      MsgBox "Descri��o � Obrigat�rio!", vbInformation, "Aten��o."
      txt_Descricao.SetFocus
      Exit Function
   End If
   
   If cmbStatus.Text = Empty Then
      MsgBox "Status � Obrigat�rio!", vbInformation, "Aten��o."
      cmbStatus.SetFocus
      Exit Function
   End If
   
   
  '' Valida��o tipo de campos
   
   If Len(txt_NumeroCartao) <> 16 Then
      MsgBox "N�mero do Cart�o deve conter 16 d�gitos!", vbInformation, "Aten��o."
      txt_NumeroCartao.SetFocus
      Exit Function
   End If
   
   If txt_Valor.Text = 0 Then
      MsgBox "Valor de transa��o 0 n�o � permitido!", vbInformation, "Aten��o."
      txt_Valor.SetFocus
      Exit Function
   End If
   
   If Len(txt_Descricao.Text) > 255 Then
      MsgBox "Descri��o m�ximo permitido at� 255 caracteres!", vbInformation, "Aten��o."
      txt_Descricao.SetFocus
      Exit Function
   End If
   
fInconsistencias = False

End Function

Private Function fLimparCampos()

    txt_Data_Cadastro = Format(CDate(Now), "dd/mm/yyyy hh:nn:ss")
    txt_CodTransacao.Text = ""
    txt_NumeroCartao.Text = Empty
    txt_Valor.Text = 0
    txt_Valor_LostFocus
    txt_Descricao.Text = ""
    cmbStatus.ListIndex = 0
    
End Function

Private Function fGravar()
Dim rs As New ADODB.Recordset
Dim SQL As String

On Error GoTo TratamentoDeErro

        '' Rotina de Gravar realiza:
        '' Gravar (id_transacao) autonumerico
        '' Alterar caso (id_transacao) diferente de 0(zero)
        '' Status Aprovado n�o permite quaisquer altera��o
        

        SQL = "select * from tb_Transacoes where Id_Transacao = " & Wval(txt_CodTransacao)
        rs.Open SQL, cn, 3, 3
        If rs.EOF Then
            rs.AddNew
            rs!Numero_Cartao = Mid(txt_NumeroCartao.Text, 1, 16) & ""
            rs!Data_Transacao = Format(CDate(Now), "dd/mm/yyyy hh:nn:ss")
            rs!Valor_Transacao = (txt_Valor.Text) & ""
            rs!Descricao = Mid(txt_Descricao.Text, 1, 255) & ""
            rs!Status = cmbStatus.ListIndex + 1
            rs.Update
            rs.Close
            Set rs = Nothing
            
            If Wval(txt_CodTransacao) = 0 Then
            SQL = "SELECT @@IDENTITY AS 'Id_Transacao'"
            rs.Open SQL, cn, 3, adLockReadOnly
                If Not rs.EOF Then
                    txt_CodTransacao = (rs!Id_Transacao)
                End If
                MsgBox "Registro salvo com sucesso! ", vbInformation, "Aten��o."
            Else
                MsgBox "Registro alterado com sucesso! ", vbInformation, "Aten��o."
            End If
            rs.Close
            Set rs = Nothing
        ElseIf Not rs.EOF Then
        
            If rs!Status <> 1 Then
                rs!Numero_Cartao = Mid(txt_NumeroCartao.Text, 1, 16) & ""
                rs!Data_Transacao = Format(CDate(Now), "dd/mm/yyyy hh:nn:ss")
                rs!Valor_Transacao = (txt_Valor.Text) & ""
                rs!Descricao = Mid(txt_Descricao.Text, 1, 255) & ""
                rs!Status = cmbStatus.ListIndex + 1
                rs.Update
                rs.Close
                Set rs = Nothing
                
                If Wval(txt_CodTransacao) = 0 Then
                SQL = "SELECT @@IDENTITY AS 'Id_Transacao'"
                rs.Open SQL, cn, 3, adLockReadOnly
                    If Not rs.EOF Then
                        txt_CodTransacao = (rs!Id_Transacao)
                    End If
                    MsgBox "Registro salvo com sucesso! ", vbInformation, "Aten��o."
                Else
                    MsgBox "Registro alterado com sucesso! ", vbInformation, "Aten��o."
                End If
            Else
                MsgBox "N�o � permitido altera��o pois Status j� est� Aprovado! ", vbInformation, "Aten��o."
                rs.Close
                Set rs = Nothing
            
            End If
        End If

Exit Function
TratamentoDeErro:
    '' Monta a mensagem de log com detalhes do erro
    Dim strErroDetails As String
    strErroDetails = "Erro na rotina MinhaRotinaQuePodeGerarErro - " & _
                     "N�mero: " & Err.Number & " | " & _
                     "Descri��o: " & Err.Description & " | " & _
                     "Fonte: " & Err.Source & " | " & _
                     "�ltimaDLL: " & Err.HelpFile & " | " & _
                     "Contexto: Linha do erro/Estado da aplica��o" ' Adicione contexto se poss�vel

    '' Chama a rotina de log do m�dulo1
    Call EscreverLogErro(strErroDetails)

    '' Opcional: Avisar o usu�rio de forma amig�vel (sem mostrar detalhes t�cnicos)
    MsgBox "Ocorreu um erro inesperado. O problema foi registrado e ser� investigado.", vbCritical, "Erro"
    
End Function

Private Function fExcluir()
Dim rs As New ADODB.Recordset
Dim SQL As String

    '' Exclus�o permitido somente se Id_transacao for diferente de 0(zero)
    '' Exclus�o permitido com confirma��o do usu�rio

If Wval(txt_CodTransacao) <> 0 Then
    If MsgBox("Deseja realmente excluir a Transa��o ?", vbYesNo, "Aten��o.") = vbYes Then
        SQL = "Delete from tb_Transacoes where Id_Transacao = " & Wval(txt_CodTransacao)
        cn.Execute (SQL)
        MsgBox "Registro exclu�do com sucesso! ", vbInformation, "Aten��o."
        Call fLimparCampos
    End If
Else
    MsgBox "Selecione uma transa��o para ser exclu�do! ", vbCritical, "Aten��o."
End If

End Function

Private Sub txt_CodTransacao_KeyPress(KeyAscii As Integer)

   '' Trata campo somente n�meros
   
   If Index = 0 Then
      If KeyAscii <> 8 And KeyAscii <> 46 Then
          If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
      End If
   End If
   
End Sub

Private Sub txt_CodTransacao_LostFocus()

On Error GoTo TratamentoDeErro

    Call fCarregarDados(Wval(txt_CodTransacao))
    
Exit Sub
TratamentoDeErro:
    '' Monta a mensagem de log com detalhes do erro
    Dim strErroDetails As String
    strErroDetails = "Erro na rotina MinhaRotinaQuePodeGerarErro - " & _
                     "N�mero: " & Err.Number & " | " & _
                     "Descri��o: " & Err.Description & " | " & _
                     "Fonte: " & Err.Source & " | " & _
                     "�ltimaDLL: " & Err.HelpFile & " | " & _
                     "Contexto: Linha do erro/Estado da aplica��o" ' Adicione contexto se poss�vel

    '' Chama a rotina de log do m�dulo1
    Call EscreverLogErro(strErroDetails)

    '' Opcional: Avisar o usu�rio de forma amig�vel (sem mostrar detalhes t�cnicos)
    MsgBox "Ocorreu um erro inesperado. O problema foi registrado e ser� investigado.", vbCritical, "Erro"


End Sub

Private Sub txt_NumeroCartao_KeyPress(KeyAscii As Integer)

    '' Permite apenas n�meros (0-9), Backspace (8) e Delete (127 - raramente usado no KeyPress)
    
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0 ' Cancela a entrada do caractere
    End If
    
End Sub

Private Sub txt_Valor_KeyPress(KeyAscii As Integer)

    '' Permite apenas n�meros (0-9), Backspace (para apagar) e o separador decimal.
    '' Em sistemas com configura��o regional do Brasil, o separador decimal � a V�RGULA (c�digo ASCII 44).
    '' Se seu sistema usa PONTO, altere 44 para 46.

    Const ASCII_VIRGULA As Integer = 44 ' C�digo ASCII para a v�rgula
    Const ASCII_PONTO As Integer = 46    ' C�digo ASCII para o ponto

    '' Condi��o para permitir:
    '' 1. D�gitos de 0 a 9
    '' 2. Tecla Backspace (vbKeyBack)
    '' 3. O separador decimal configurado (v�rgula ou ponto)
    
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) And _
        KeyAscii <> vbKeyBack And _
        KeyAscii <> ASCII_VIRGULA Then ' Use ASCII_PONTO se o separador for ponto
        KeyAscii = 0 ' Cancela a entrada do caractere se n�o for permitido
    End If

    '' Garante que apenas um separador decimal seja digitado
    If KeyAscii = ASCII_VIRGULA Then ' Use ASCII_PONTO se o separador for ponto
        If InStr(txt_Valor.Text, ",") > 0 Then ' Verifica se j� existe uma v�rgula no texto
            KeyAscii = 0 ' Cancela a entrada da v�rgula se j� houver uma
        End If
    End If
    
End Sub

Private Sub txt_Valor_LostFocus()
    Dim dblValor As Double

    '' 1. Verifica se o campo n�o est� vazio
    If Len(txt_Valor.Text) > 0 Then

        '' 2. Checa se o conte�do � um n�mero v�lido.
        '' IsNumeric considera a configura��o regional do seu Windows (v�rgula ou ponto como decimal).
        If IsNumeric(txt_Valor.Text) Then
            '' 3. Converte o texto para um n�mero decimal (Double)
            dblValor = CDbl(txt_Valor.Text)

            '' 4. Formata o n�mero com 2 casas decimais e atualiza o campo.
            ''    FormatNumber aplica o separador de milhares e decimais conforme a regi�o.
            txt_Valor.Text = FormatNumber(dblValor, 2) ' O "2" define 2 casas decimais
        Else
            ' '5. Se n�o for um n�mero v�lido, avisa o usu�rio e limpa o campo.
            MsgBox "Valor inv�lido. Por favor, digite um n�mero.", vbExclamation, "Erro de Entrada"
            txt_Valor.Text = "" ' Limpa o campo
            txt_Valor.SetFocus  ' Devolve o foco para o campo para o usu�rio corrigir
        End If
    End If

End Sub

Private Function fCarregarDados(nCod As Integer)
Dim rs As New ADODB.Recordset
Dim SQL As String

    '' Carrega os Dados das Transa��es
    SQL = "select * from tb_Transacoes where id_transacao = " & Wval(nCod)
    rs.Open SQL, cn, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        txt_Data_Cadastro = Format(CDate(rs!Data_Transacao), "dd/mm/yyyy hh:nn:ss")
        txt_Descricao.Text = rs!Descricao & ""
    
        If (rs!Status = "1") Then
            cmbStatus.ListIndex = 0
        ElseIf (rs!Status = "2") Then
            cmbStatus.ListIndex = 1
        Else
            cmbStatus.ListIndex = 2
        End If
        txt_NumeroCartao = Trim(Str(rs!Numero_Cartao))
        txt_Valor = Wval(rs!Valor_Transacao)
    
    Else
        MsgBox "Nenhum registro foi Encontrado! Verifique!", vbInformation, "Aten��o."
        Call fLimparCampos
        Exit Function
    End If

End Function
