VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_Consulta 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Transa��es"
   ClientHeight    =   6285
   ClientLeft      =   3750
   ClientTop       =   3195
   ClientWidth     =   7710
   Icon            =   "frm_Consulta.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7710
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   1164
      ButtonWidth     =   1085
      ButtonHeight    =   1005
      Appearance      =   1
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
      Left            =   5160
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txt_CodTransacao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   6975
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         ItemData        =   "frm_Consulta.frx":0A4E
         Left            =   3240
         List            =   "frm_Consulta.frx":0A5B
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txt_Descricao 
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
         Height          =   1335
         IMEMode         =   3  'DISABLE
         Left            =   3240
         MaxLength       =   255
         TabIndex        =   3
         ToolTipText     =   "at� 255 Caracteres."
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txt_Valor 
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
         TabIndex        =   2
         ToolTipText     =   "Somente decimais positivos"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txt_NumeroCartao 
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
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "* Descri��o :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1920
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "* Valor :"
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
         TabIndex        =   10
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "* N�mero Cart�o :"
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
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.PictureBox ImageList1 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   12360
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   13
      Top             =   3120
      Width           =   1200
   End
   Begin VB.Label Label6 
      Caption         =   "* N�mero Cart�o :"
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
      Left            =   360
      TabIndex        =   15
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Transa��o :"
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
      Left            =   3480
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Id.Transa��o:"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "frm_Consulta"
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
    Me.Caption = "XYZ - Administradora de Cart�es de Cr�dito - " + Me.Caption
    Call fLimparCampos
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
    ' Monta a mensagem de log com detalhes do erro
    Dim strErroDetails As String
    strErroDetails = "Erro na rotina MinhaRotinaQuePodeGerarErro - " & _
                     "N�mero: " & Err.Number & " | " & _
                     "Descri��o: " & Err.Description & " | " & _
                     "Fonte: " & Err.Source & " | " & _
                     "�ltimaDLL: " & Err.HelpFile & " | " & _
                     "Contexto: Linha do erro/Estado da aplica��o" ' Adicione contexto se poss�vel

    ' Chama a rotina de log do m�dulo1
    Call EscreverLogErro(strErroDetails)

    ' Opcional: Avisar o usu�rio de forma amig�vel (sem mostrar detalhes t�cnicos)
    MsgBox "Ocorreu um erro inesperado. O problema foi registrado e ser� investigado.", vbCritical, "Erro"


End Sub

Private Function fInconsistencias()

fInconsistencias = True


'Valida campos obrigat�rios

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
   
   
  'Valida��o tipo de campos
   
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
    txt_NumeroCartao.Text = ""
    txt_Valor.Text = ""
    txt_Descricao.Text = ""
    cmbStatus.ListIndex = 0
    


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

    'estamos usando a propriedade Clip e o m�todo GetString para selecionar uma regi�o do grid
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

    .FormatString = "||>C�digo.|<Nome|^Dt.Cadastro|<Usuario|>Senha|>Adm"

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

On Error GoTo TratamentoDeErro


        SQL = "select * from tb_Transacoes where Id_Transacao = " & Wval(txt_CodTransacao)
        rs.Open SQL, cn, 3, 3
        If rs.EOF Then
            rs.AddNew
            rs!Numero_Cartao = Mid(txt_NumeroCartao.Text, 1, 16) & ""
            rs!Data_Transacao = Format(CDate(Now), "dd/mm/yyyy hh:nn:ss")
            rs!Valor_Transacao = (txt_Valor.Text) & ""
            rs!Descricao = Mid(txt_Descricao.Text, 1, 255) & ""
            rs!Status = cmbStatus.ListIndex
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
        
            If rs!Status <> 0 Then
                rs!Numero_Cartao = Mid(txt_NumeroCartao.Text, 1, 16) & ""
                rs!Data_Transacao = Format(CDate(Now), "dd/mm/yyyy hh:nn:ss")
                rs!Valor_Transacao = (txt_Valor.Text) & ""
                rs!Descricao = Mid(txt_Descricao.Text, 1, 255) & ""
                rs!Status = cmbStatus.ListIndex
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
    ' Monta a mensagem de log com detalhes do erro
    Dim strErroDetails As String
    strErroDetails = "Erro na rotina MinhaRotinaQuePodeGerarErro - " & _
                     "N�mero: " & Err.Number & " | " & _
                     "Descri��o: " & Err.Description & " | " & _
                     "Fonte: " & Err.Source & " | " & _
                     "�ltimaDLL: " & Err.HelpFile & " | " & _
                     "Contexto: Linha do erro/Estado da aplica��o" ' Adicione contexto se poss�vel

    ' Chama a rotina de log do m�dulo1
    Call EscreverLogErro(strErroDetails)

    ' Opcional: Avisar o usu�rio de forma amig�vel (sem mostrar detalhes t�cnicos)
    MsgBox "Ocorreu um erro inesperado. O problema foi registrado e ser� investigado.", vbCritical, "Erro"


'Call fFormatarGradeUsuarios
'Call fCarregaDadosGrade
    
End Function

Private Function fExcluir()
Dim rs As New ADODB.Recordset
Dim SQL As String

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


Private Sub txt_NumeroCartao_KeyPress(KeyAscii As Integer)
 ' Permite apenas n�meros (0-9), Backspace (8) e Delete (127 - raramente usado no KeyPress)
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0 ' Cancela a entrada do caractere
    End If
End Sub

Private Sub txt_Valor_KeyPress(KeyAscii As Integer)
 ' Permite apenas n�meros (0-9), Backspace (para apagar) e o separador decimal.
    ' Em sistemas com configura��o regional do Brasil, o separador decimal � a V�RGULA (c�digo ASCII 44).
    ' Se seu sistema usa PONTO, altere 44 para 46.

    Const ASCII_VIRGULA As Integer = 44 ' C�digo ASCII para a v�rgula
    Const ASCII_PONTO As Integer = 46    ' C�digo ASCII para o ponto

    ' Condi��o para permitir:
    ' 1. D�gitos de 0 a 9
    ' 2. Tecla Backspace (vbKeyBack)
    ' 3. O separador decimal configurado (v�rgula ou ponto)
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) And _
       KeyAscii <> vbKeyBack And _
       KeyAscii <> ASCII_VIRGULA Then ' Use ASCII_PONTO se o separador for ponto

        KeyAscii = 0 ' Cancela a entrada do caractere se n�o for permitido
    End If

    ' Garante que apenas um separador decimal seja digitado
    If KeyAscii = ASCII_VIRGULA Then ' Use ASCII_PONTO se o separador for ponto
        If InStr(txt_Valor.Text, ",") > 0 Then ' Verifica se j� existe uma v�rgula no texto
            KeyAscii = 0 ' Cancela a entrada da v�rgula se j� houver uma
        End If
    End If
End Sub

Private Sub txt_Valor_LostFocus()
    Dim dblValor As Double

    ' 1. Verifica se o campo n�o est� vazio
    If Len(txt_Valor.Text) > 0 Then

        ' 2. Checa se o conte�do � um n�mero v�lido.
        '    IsNumeric considera a configura��o regional do seu Windows (v�rgula ou ponto como decimal).
        If IsNumeric(txt_Valor.Text) Then
            ' 3. Converte o texto para um n�mero decimal (Double)
            dblValor = CDbl(txt_Valor.Text)

            ' 4. Formata o n�mero com 2 casas decimais e atualiza o campo.
            '    FormatNumber aplica o separador de milhares e decimais conforme a regi�o.
            txt_Valor.Text = FormatNumber(dblValor, 2) ' O "2" define 2 casas decimais
        Else
            ' 5. Se n�o for um n�mero v�lido, avisa o usu�rio e limpa o campo.
            MsgBox "Valor inv�lido. Por favor, digite um n�mero.", vbExclamation, "Erro de Entrada"
            txt_Valor.Text = "" ' Limpa o campo
            txt_Valor.SetFocus  ' Devolve o foco para o campo para o usu�rio corrigir
        End If
    End If

End Sub
