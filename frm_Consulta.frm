VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_Consulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Transações"
   ClientHeight    =   6765
   ClientLeft      =   3750
   ClientTop       =   3195
   ClientWidth     =   7620
   Icon            =   "frm_Consulta.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7620
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   1588
      ButtonWidth     =   1508
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pesquisar"
            Key             =   "Pesquisar"
            Object.ToolTipText     =   "Pesquisar Transações"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fechar"
            Key             =   "Fechar"
            Object.ToolTipText     =   "Fechar"
            Object.Tag             =   ""
            ImageIndex      =   2
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
      Picture         =   "frm_Consulta.frx":1084A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Buscar Transações"
      Top             =   1080
      Visible         =   0   'False
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
         Height          =   2325
         Left            =   3240
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         ToolTipText     =   "até 255 Caracteres."
         Top             =   1200
         Width           =   3015
      End
      Begin VB.ComboBox cmbStatus 
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
         ItemData        =   "frm_Consulta.frx":10A7D
         Left            =   3240
         List            =   "frm_Consulta.frx":10A8A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3720
         Width           =   2775
      End
      Begin VB.TextBox txt_Valor 
         Appearance      =   0  'Flat
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
         Left            =   3240
         TabIndex        =   2
         ToolTipText     =   "Somente decimais positivos"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txt_NumeroCartao 
         Appearance      =   0  'Flat
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
         Left            =   3240
         MaxLength       =   16
         TabIndex        =   1
         ToolTipText     =   "16 dígitos"
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
         TabIndex        =   12
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "* Descrição :"
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
         TabIndex        =   10
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
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "* Número Cartão :"
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
   Begin VB.PictureBox ImageList1 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   12360
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   14
      Top             =   3120
      Width           =   1200
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   9240
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Consulta.frx":10AB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Consulta.frx":111C7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Transação:"
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
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Id.Transação:"
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
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "frm_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Buscar_Click()

On Error GoTo TratamentoDeErro
    
    '' Título DataGrid
    frmLocalizar.lstTitulo.List(0) = "Id_Transacao"
    frmLocalizar.lstTitulo.List(1) = "Numero_Cartao"
    frmLocalizar.lstTitulo.List(2) = "Data_Transacao"
    frmLocalizar.lstTitulo.List(3) = "Status"
    frmLocalizar.lstTitulo.List(4) = "Valor_Transacao"
    frmLocalizar.lstTitulo.List(5) = "Descricao"
    
    '' Título BD
    frmLocalizar.lstCampo.List(0) = "Id_Transacao"
    frmLocalizar.lstCampo.List(1) = "Numero_Cartao"
    frmLocalizar.lstCampo.List(2) = "CONVERT(VARCHAR(10), Data_Transacao, 103)"
    frmLocalizar.lstCampo.List(3) = "CASE WHEN Status = 1 THEN 'Aprovada' WHEN Status = 2 THEN 'Pendente' ELSE 'Cancelada' end"
    frmLocalizar.lstCampo.List(4) = "replace(CONVERT(NUMERIC(18, 2), Valor_Transacao ),'.',',')"
    frmLocalizar.lstCampo.List(5) = "Descricao"
    
    '' Parâmetro da Tabela
    frmLocalizar.txtTabela.Text = "tb_Transacoes"
    frmLocalizar.lstOrdem.List(0) = "Id_Transacao"
    frmLocalizar.lstRetorna.List(0) = "0"
    frmLocalizar.FMontaChave
    frmLocalizar.Caption = " Buscar Transações"
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
                     "Número: " & Err.Number & " | " & _
                     "Descrição: " & Err.Description & " | " & _
                     "Fonte: " & Err.Source & " | " & _
                     "ÚltimaDLL: " & Err.HelpFile & " | " & _
                     "Contexto: Linha do erro/Estado da aplicação" ' Adicione contexto se possível

    '' Chama a rotina de log do módulo1
    Call EscreverLogErro(strErroDetails)

    '' Opcional: Avisar o usuário de forma amigável (sem mostrar detalhes técnicos)
    MsgBox "Ocorreu um erro inesperado. O problema foi registrado e será investigado.", vbCritical, "Erro"

End Sub

Private Sub Form_Load()

On Error GoTo TratamentoDeErro

    '' Renomeia Título da janela e Limpa os Campos
    Me.Caption = "XYZ - Administradora de Cartões de Crédito - " + Me.Caption
    Call fLimparCampos
    
    
Exit Sub
TratamentoDeErro:
    '' Monta a mensagem de log com detalhes do erro
    Dim strErroDetails As String
    strErroDetails = "Erro na rotina MinhaRotinaQuePodeGerarErro - " & _
                     "Número: " & Err.Number & " | " & _
                     "Descrição: " & Err.Description & " | " & _
                     "Fonte: " & Err.Source & " | " & _
                     "ÚltimaDLL: " & Err.HelpFile & " | " & _
                     "Contexto: Linha do erro/Estado da aplicação" ' Adicione contexto se possível

    '' Chama a rotina de log do módulo1
    Call EscreverLogErro(strErroDetails)

    '' Opcional: Avisar o usuário de forma amigável (sem mostrar detalhes técnicos)
    MsgBox "Ocorreu um erro inesperado. O problema foi registrado e será investigado.", vbCritical, "Erro"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

On Error GoTo TratamentoDeErro

    Select Case Button.Key
    
        Case "Pesquisar"
            Call cmd_Buscar_Click
         
         
        Case "Fechar"
            Unload Me
    
End Select

Exit Sub
TratamentoDeErro:
    '' Monta a mensagem de log com detalhes do erro
    Dim strErroDetails As String
    strErroDetails = "Erro na rotina MinhaRotinaQuePodeGerarErro - " & _
                     "Número: " & Err.Number & " | " & _
                     "Descrição: " & Err.Description & " | " & _
                     "Fonte: " & Err.Source & " | " & _
                     "ÚltimaDLL: " & Err.HelpFile & " | " & _
                     "Contexto: Linha do erro/Estado da aplicação" ' Adicione contexto se possível

    '' Chama a rotina de log do módulo1
    Call EscreverLogErro(strErroDetails)

    '' Opcional: Avisar o usuário de forma amigável (sem mostrar detalhes técnicos)
    MsgBox "Ocorreu um erro inesperado. O problema foi registrado e será investigado.", vbCritical, "Erro"


End Sub


Private Function fLimparCampos()

    txt_Data_Cadastro = Format(CDate(Now), "dd/mm/yyyy hh:nn:ss")
    txt_CodTransacao.Text = ""
    txt_NumeroCartao.Text = Empty
    txt_Valor.Text = 0
    txt_Descricao.Text = ""
    cmbStatus.ListIndex = 0
    
End Function


Private Sub txt_CodTransacao_LostFocus()

On Error GoTo TratamentoDeErro

    Call fCarregarDados(Wval(txt_CodTransacao))
    
    
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

Private Function fCarregarDados(nCod As Integer)
Dim rs As New ADODB.Recordset
Dim SQL As String

    '' Carrega os Dados das Transações
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
        MsgBox "Nenhum registro foi Encontrado! Verifique!", vbInformation, "Atenção."
        Call fLimparCampos
        Exit Function
    End If


End Function
