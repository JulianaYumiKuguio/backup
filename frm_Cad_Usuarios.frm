VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_Cad_Usuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Usuários / Perfil de Acesso"
   ClientHeight    =   8805
   ClientLeft      =   3750
   ClientTop       =   3195
   ClientWidth     =   6555
   Icon            =   "frm_Cad_Usuarios.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   6555
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   1588
      ButtonWidth     =   1085
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
   Begin VB.TextBox txt_Data_Cadastro 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox txt_Cod 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   240
      TabIndex        =   6
      Top             =   6360
      Width           =   6135
      Begin VB.CheckBox chk_Adm 
         Caption         =   "Administrador"
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txt_Senha 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3240
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "Mínimo 6, Máximo 10 Caracteres."
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txt_Usuario 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         MaxLength       =   100
         TabIndex        =   2
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txt_Nome 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "* Senha :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "* Usuário :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "* Nome :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Cadastro :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6120
      Width           =   1335
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   12360
      Top             =   3120
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
            Picture         =   "frm_Cad_Usuarios.frx":0A4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cad_Usuarios.frx":1964
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cad_Usuarios.frx":21B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cad_Usuarios.frx":2F88
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cad_Usuarios.frx":3CDA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_Cad_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub flxUsuarios_Click()
If flxUsuarios.Rows > 1 And Wval(flxUsuarios.TextMatrix(flxUsuarios.RowSel, 2)) > 0 Then
    txt_Cod.Text = Wval(flxUsuarios.TextMatrix(flxUsuarios.RowSel, 2))
    txt_Nome = flxUsuarios.TextMatrix(flxUsuarios.RowSel, 3) & ""
    txt_Data_Cadastro = Format(flxUsuarios.TextMatrix(flxUsuarios.RowSel, 4), "dd/mm/yyyy") & ""
    txt_Usuario = flxUsuarios.TextMatrix(flxUsuarios.RowSel, 5) & ""
    txt_Senha = flxUsuarios.TextMatrix(flxUsuarios.RowSel, 6) & ""
    If Wval(flxUsuarios.TextMatrix(flxUsuarios.RowSel, 7)) = 1 Then
        chk_Adm.Value = 1
    Else
        chk_Adm.Value = 0
    End If
End If
End Sub

Private Sub Form_Load()
    Me.Caption = "ATB Info - " + Me.Caption
    Call fLimparCampos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
    Case "Novo"
        Call fLimparCampos
     
    Case "Salvar"
        If nAdministrador = True Then
            If Not fInconsistencias Then
                Call fGravar
            End If
        Else
            MsgBox "Você não possui privilégios de Administrador do Sistema!", vbInformation, "Atenção."
        End If
    Case "Excluir"
        If nAdministrador = True Then
            If (txt_Cod) > 0 Then
                Call fExcluir
            Else
                MsgBox "É necessário um Registro para Exclusão!", vbInformation, "Atenção."
            End If
        Else
            MsgBox "Você não possui privilégios de Administrador do Sistema!", vbInformation, "Atenção."
        End If
        
        Case "Fechar"
        Unload Me
    
End Select
End Sub

Private Function fInconsistencias()

fInconsistencias = True

   If txt_Nome.Text = Empty Then
      MsgBox "Nome é Obrigatório!", vbInformation, "Atenção."
      txt_Nome.SetFocus
      Exit Function
   End If
   
   If txt_Usuario.Text = Empty Then
      MsgBox "Usuário é Obrigatório!", vbInformation, "Atenção."
      txt_Usuario.SetFocus
      Exit Function
   End If
   
   If txt_Senha.Text = Empty Then
      MsgBox "Senha é Obrigatório!", vbInformation, "Atenção."
      txt_Senha.SetFocus
      Exit Function
   End If
   
   If Len(txt_Senha) < 6 Or Len(txt_Senha) > 10 Then
      MsgBox "Senha deve conter no Mínimo 6 caracteres e no Máximo 10 caracteres!", vbInformation, "Atenção."
      txt_Senha.SetFocus
      Exit Function
   End If
   
fInconsistencias = False

End Function

Private Function fLimparCampos()
    txt_Data_Cadastro = Format(CDate(Now), "dd/mm/yyyy")
    txt_Cod = ""
    txt_Nome = ""
    txt_Senha = ""
    txt_Usuario = ""
    chk_Adm.Value = False

    'Call fFormatarGradeUsuarios
    'Call fCarregaDadosGrade
End Function


Private Function fCarregaDadosGrade()
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim i As Integer

Dim lTimer As Long

Screen.MousePointer = vbHourglass

flxUsuarios.Refresh
lTimer = Timer

flxUsuarios.Visible = False


rs.Open "Select '' as sel1,'' as sel2, codigo as Codigo, Nome as Nome,Data_Cadastro as Data_Cadastro, Usuario as Usuario,Senha as Senha,Adm as Adm  from tb_Usuarios ORDER BY Codigo asc", cn, 3, 3
If Not rs.EOF Then
    rs.MoveFirst

    'define o numero de linhas e colunas e configura o grid
    flxUsuarios.Rows = rs.RecordCount + 1
    flxUsuarios.Row = 1
    flxUsuarios.Col = 0
    flxUsuarios.RowSel = flxUsuarios.Rows - 1
    flxUsuarios.ColSel = flxUsuarios.Cols - 1

    'estamos usando a propriedade Clip e o método GetString para selecionar uma região do grid
    flxUsuarios.Clip = rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
    flxUsuarios.Row = 1
    flxUsuarios.Visible = True
End If
'libera os objetos
Set rs = Nothing
Set db = Nothing

Screen.MousePointer = vbDefault

End Function

Private Function fFormatarGradeUsuarios()
Dim i As Integer

With flxUsuarios
    .Clear
    .Rows = 2
    .Cols = 8

    .FormatString = "||>Código.|<Nome|^Dt.Cadastro|<Usuario|>Senha|>Adm"

    For i = 0 To .Cols - 1
        .Row = 0
        .Col = i
        .CellFontBold = True
    Next i
 .ColWidth(0) = 200
 .ColWidth(1) = 0
 .ColWidth(2) = 900
 .ColWidth(3) = 2500
 .ColWidth(4) = 0
 .ColWidth(5) = 2500
 .ColWidth(6) = 0
 .ColWidth(7) = 0
End With
End Function


Private Function fGravar()
Dim rs As New ADODB.Recordset
Dim SQL As String

SQL = "select * from tb_Usuarios where Codigo = " & Wval(txt_Cod)
rs.Open SQL, cn, 3, 3
If rs.EOF Then
    rs.AddNew
End If

    rs!Nome = Mid(txt_Nome.Text, 1, 100) & ""
    rs!Data_Cadastro = Format(CDate(Now), "dd/mm/yyyy")
    rs!Usuario = Mid(txt_Usuario.Text, 1, 100) & ""
    rs!Senha = Mid(txt_Senha.Text, 1, 10) & ""
    rs!Adm = IIf(chk_Adm, 1, 0)
    rs.Update
rs.Close
Set rs = Nothing
If Wval(txt_Cod) = 0 Then
    SQL = "SELECT @@IDENTITY AS 'cod'"
    rs.Open SQL, cn, 3, adLockReadOnly
    If Not rs.EOF Then
        txt_Cod = Wval(rs!cod)
    End If
End If
MsgBox "Registro salvo com sucesso! ", vbInformation, "Atenção."

'Call fFormatarGradeUsuarios
'Call fCarregaDadosGrade
    
End Function

Private Function fExcluir()
Dim rs As New ADODB.Recordset
Dim SQL As String

If Wval(txt_Cod) <> 0 Then
    If MsgBox("Deseja realmente excluir o registro ?", vbYesNo, "Atenção.") = vbYes Then
        SQL = "Delete from tb_Usuarios where codigo = " & Wval(txt_Cod)
        cn.Execute (SQL)
        MsgBox "Registro excluído com sucesso! ", vbInformation, "Atenção."
        Call fLimparCampos
    End If
Else
    MsgBox "Selecione um registro para ser excluído! ", vbCritical, "Atenção."
End If

End Function


