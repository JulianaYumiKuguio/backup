VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLocalizar 
   Caption         =   " Localizar"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13335
   Icon            =   "frmLocalizar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   13335
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tobLocalizar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imgItens"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Fechar"
            Object.ToolTipText     =   "Fechar"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "OrdenarAZ"
            Object.ToolTipText     =   "Ordenar A - Z"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "OrdenarZA"
            Object.ToolTipText     =   "Ordenar Z - A"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Filtrar"
            Object.ToolTipText     =   "Filtrar registros"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Localiza"
            Object.ToolTipText     =   "Localizar registro"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Primeiro"
            Object.ToolTipText     =   "Vai para o primeiro registro"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Anterior"
            Object.ToolTipText     =   "Vai para o registro anterior"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Proximo"
            Object.ToolTipText     =   "Vai para o próximo registro"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Vai para o último registro"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.TextBox TxtRecord 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   7560
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "frmLocalizar.frx":1084A
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Align           =   2  'Align Bottom
      Height          =   7515
      Left            =   0
      TabIndex        =   26
      Top             =   1335
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   13256
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   24
      WrapCellPointer =   -1  'True
      RowDividerStyle =   6
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Localizar"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox spnFil 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   13305
      TabIndex        =   19
      Top             =   420
      Width           =   13335
      Begin VB.TextBox txtconteudo 
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
         Height          =   405
         Left            =   3720
         TabIndex        =   27
         Top             =   360
         Width           =   4815
      End
      Begin VB.ComboBox cmbColuna 
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
         Height          =   420
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   390
         Width           =   2265
      End
      Begin VB.ComboBox cmbCondicao 
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
         Height          =   420
         Left            =   2535
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   390
         Width           =   975
      End
      Begin VB.CommandButton cmdFiltra 
         Caption         =   "Filtrar"
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
         Height          =   405
         Left            =   8625
         TabIndex        =   20
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Filtrar Coluna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   25
         Top             =   45
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Condição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2445
         TabIndex        =   24
         Top             =   45
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Conteúdo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3675
         TabIndex        =   23
         Top             =   45
         Width           =   1050
      End
   End
   Begin VB.Data dtaDados 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.ListBox lstCampo 
      Height          =   645
      Left            =   2070
      TabIndex        =   14
      Top             =   5175
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ListBox lstOrdem 
      Height          =   645
      ItemData        =   "frmLocalizar.frx":1085E
      Left            =   4950
      List            =   "frmLocalizar.frx":10860
      TabIndex        =   13
      Top             =   5175
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox lstTitulo 
      Height          =   645
      ItemData        =   "frmLocalizar.frx":10862
      Left            =   120
      List            =   "frmLocalizar.frx":10864
      TabIndex        =   12
      Top             =   7200
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox txtCondicao 
      Height          =   285
      Left            =   7065
      TabIndex        =   11
      Top             =   5175
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtTipo 
      Height          =   285
      Left            =   7065
      TabIndex        =   10
      Top             =   5490
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCondicaoUsu 
      Height          =   285
      Left            =   7695
      TabIndex        =   9
      Top             =   5490
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtTabela 
      Height          =   285
      Left            =   6435
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtBanco 
      Height          =   285
      Left            =   6435
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ListBox lstRetorna 
      Height          =   645
      Left            =   8280
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtTabela1 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtTabela2 
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox TxtTabela3 
      Height          =   285
      Left            =   3240
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox TxtRestrito 
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   5880
      Width           =   495
   End
   Begin VB.PictureBox spnLoc 
      BackColor       =   &H00C0C0C0&
      Height          =   690
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   9615
      TabIndex        =   15
      Top             =   420
      Visible         =   0   'False
      Width           =   9675
      Begin VB.TextBox txtPalavra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Top             =   315
         Width           =   4020
      End
      Begin VB.CommandButton cmdFechar 
         Caption         =   "&Fechar"
         Height          =   285
         Left            =   8400
         TabIndex        =   16
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Digite as primeiras letras da palavra que está procurando"
         Height          =   195
         Left            =   2250
         TabIndex        =   18
         Top             =   90
         Width           =   4005
      End
   End
   Begin ComctlLib.ImageList imgItens 
      Left            =   135
      Top             =   5175
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLocalizar.frx":10866
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLocalizar.frx":10A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLocalizar.frx":10B7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLocalizar.frx":10C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLocalizar.frx":10DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLocalizar.frx":10EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLocalizar.frx":10FD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLocalizar.frx":110E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLocalizar.frx":111FA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLocalizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wBuscaRetorna As String
Dim wCampo As Integer
Dim i As Byte
Function FCarregaOpcoesFiltro()

    With cmbColuna
        .Clear
        For i = 0 To lstTitulo.ListCount
            .AddItem lstTitulo.List(i)
        Next i
        '.SetFocus
    End With
    
    With cmbCondicao
        .Clear
        .AddItem "="
        .AddItem "<>"
        .AddItem "<"
        .AddItem ">"
        .AddItem ">="
        .AddItem "<="
    End With
    
    txtconteudo.Text = Empty

End Function


Function FPosNormalFil()

'    dbgDados.Height = 4230
'    dbgDados.Left = 45
'    dbgDados.Top = 90
    
'    spnGrade.Height = 4380
'    spnGrade.Left = 45
'    spnGrade.Top = 450

End Function

Function FPosLocalizar()

'    DataGrid1.Height = 3465
'    DataGrid1.Left = 45
'    DataGrid1.Top = 90
'
'    spnGrade.Height = 3615
'    spnGrade.Left = 45
'    spnGrade.Top = 1215
'
'    spnLoc.Visible = True

End Function

Function FPosFiltrar()

'    DataGrid1.Height = 3465
'    DataGrid1.Left = 45
'    DataGrid1.Top = 90
'
'    spnGrade.Height = 3615
'    spnGrade.Left = 45
'    spnGrade.Top = 1215

   '' spnFil.Visible = True

End Function

Function FMontaChave()
                    
    Dim sSelecao As String
    Dim i As Integer
   
    'Monta Chaves do Select;
    
    dtaDados.DatabaseName = txtBanco.Text
    
    'Monta parte dos Campos e dos Titulos;
    If TxtRestrito.Text = Empty Then
      sSelecao = "SELECT "
    Else
      sSelecao = "SELECT " & TxtRestrito.Text & " "
    End If
    
    For i = 0 To lstCampo.ListCount - 1
        If Format(lstCampo.List(i), ">") <> Format(lstTitulo.List(i), ">") And lstTitulo.List(i) <> Empty Then
            sSelecao = sSelecao & lstCampo.List(i) & " As " & lstTitulo.List(i) & ","
        Else
            sSelecao = sSelecao & lstCampo.List(i) & ","
        End If
    Next i
    sSelecao = Mid(sSelecao, 1, Len(sSelecao) - 1)
    
    'Monta Parte da Tabela;
'    sSelecao = sSelecao & " From [" & txtTabela.Text & "] "
    sSelecao = sSelecao & " From " & txtTabela.Text & " "
    
    'Monta Até treis Tabelas de um mesmo Banco de Dados
    If txtTabela1.Text <> Empty Then
        sSelecao = sSelecao & "a , [" & txtTabela1.Text & "] b"
    End If
    If txtTabela2.Text <> Empty Then
        sSelecao = sSelecao & ", [" & txtTabela2.Text & "] c"
    End If
    If TxtTabela3.Text <> Empty Then
        sSelecao = sSelecao & ", [" & TxtTabela3.Text & "] d"
    End If
    
    'Monta Parte da Condição;
    If txtCondicao.Text <> Empty Then
        sSelecao = sSelecao & " Where " & txtCondicao.Text
    End If
    
    'Verifica se Usuario selecionou Filtro Proprio;
    If txtCondicaoUsu <> Empty Then
        If txtCondicao.Text = Empty Then
        
            If InStr(1, txtCondicaoUsu.Text, "replace(CONVERT(NUMERIC(18, 2), Valor_Transacao ),'.',',')", vbTextCompare) > 0 Then
            
                txtCondicaoUsu.Text = Replace(txtCondicaoUsu.Text, "replace(CONVERT(NUMERIC(18, 2), Valor_Transacao ),'.',',')", "Valor_Transacao")
            
            End If
        
            sSelecao = sSelecao & " Where " & txtCondicaoUsu.Text
        Else
            sSelecao = sSelecao & " And " & txtCondicaoUsu.Text
        End If
    End If
    
    'Monta Parte da Ordem;
    sSelecao = sSelecao & " Order By "
    For i = 0 To lstOrdem.ListCount - 1
        sSelecao = sSelecao & lstOrdem.List(i) & " " & txtTipo.Text & ","
    Next i
    sSelecao = Mid(sSelecao, 1, Len(sSelecao) - 1) & ";"
    
'     SQL usando varias tabelas de um mesmo MDB;
'     sSelecao = "Select a.NumeroPedido,a.DataPedido,a.CodigoCliente,b.CodigoProduto From " & _
               "[Pedido Master] a, [Pedido Itens] b,[Produto] c Where b.NumeroPedido=a.NumeroPedido And a.NumeroPedido='" & variavel & "' And b.CodigoProduto=c.Codigo Order By a.NumeroPedido ASC"
    
     Dim rs As New ADODB.Recordset
     Dim SQL As String
     
     
        SQL = sSelecao
        rs.Open SQL, cn, 3, 3
            

        If Not rs.EOF Then
            Set DataGrid1.DataSource = PaginarRecordset(rs)
        End If
        rs.Close
        Set rs = Nothing
        
        
        DataGrid1.Refresh
        
''    dtaDados.RecordSource = sSelecao
''    dtaDados.Refresh
''    dbgDados.Refresh
''    TxtRecord.Text = "(" & Str(dtaDados.Recordset.RecordCount) & ")-Registros"
'''    dtaDados.Recordset.MoveFirst
''
''    ' Definir Tamanho das Colunas.
''
''    For i = 0 To (dtaDados.Recordset.Fields.Count) - 1
''        If dtaDados.Recordset.Fields(i).Type = 5 Then
''            dbgDados.Columns(i).NumberFormat = "##0.00"
''        End If
''
''        If dtaDados.Recordset.Fields(i).Type = 12 Then
''            dbgDados.Columns(i).Width = 6500
''        Else
''            If dtaDados.Recordset.Fields(i).Size > Len(dtaDados.Recordset.Fields(i).Name) Then
''                dbgDados.Columns(i).Width = dtaDados.Recordset.Fields(i).Size * 100
''            Else
''                dbgDados.Columns(i).Width = Len(dtaDados.Recordset.Fields(i).Name) * 100
''            End If
''        End If
''    Next i

 FCarregaOpcoesFiltro
End Function

Function FOrdenaColunas(sTipo As String)

    Dim i As Integer
    
    'Busca Colunas Selecionadas para Ordenar;
    If DataGrid1.SelStartCol = -1 Or DataGrid1.SelEndCol = -1 Then
        MsgBox "Selecione uma coluna a ser ordenada!", 48, "Aviso"
        Exit Function
    End If
    
    lstOrdem.Clear
    
    For i = DataGrid1.SelStartCol To DataGrid1.SelEndCol
        lstOrdem.AddItem lstCampo.List(i)
    Next i
   
    txtTipo.Text = sTipo
    
    FMontaChave

End Function


Private Sub cmbFecharFil_Click()
                
    'Tira o Filtro do Usuario;
    txtCondicaoUsu.Text = Empty
    FMontaChave
    
    tobLocalizar.Buttons(7).Value = 0
    FPosNormalFil
   ' dbgDados.SetFocus

End Sub





Private Sub cmdFechar_Click()
                
    tobLocalizar.Buttons(9).Value = 0
 
End Sub

Private Sub cmdFiltra_Click()

    Dim i As Integer
    
    'Filtro do Usuário

    If cmbColuna.Text = Empty Then
        MsgBox "Nome da Coluna a ser filtrada Obrigatória !", 48, Me.Caption
        cmbColuna.SetFocus
        Exit Sub
    End If
    
    If cmbCondicao.Text = Empty Then
        MsgBox "Condição a ser filtrada Obrigatória !", 48, Me.Caption
        cmbCondicao.SetFocus
        Exit Sub
    End If

    If txtconteudo.Text = Empty Then
        MsgBox "Conteúdo a ser filtrado Obrigatório !", 48, Me.Caption
        txtconteudo.SetFocus
        Exit Sub
    End If
    
    For i = 0 To lstTitulo.ListCount
        
        If lstTitulo.List(i) = cmbColuna.Text Then
        'DataGrid1.Text
        
            If Wval(Replace(txtconteudo.Text, """", "")) <> 0 Then
                If (VarType(Wval(Replace(txtconteudo.Text, """", "")))) = vbInteger Then
                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtconteudo
                ElseIf (VarType(Wval(Replace(txtconteudo.Text, """", "")))) = vbDouble Then
                    'txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtconteudo
                    
                    If lstCampo.List(i) = "CAST(Valor_Transacao AS VARCHAR(20))" Then
                        txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtconteudo
                        If lstCampo.List(i) = "CAST(Valor_Transacao AS VARCHAR(20))" Then lstCampo.List(i) = "Valor_Transacao"
                    Else
                         txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & "'" & txtconteudo & "'"
                    End If
                End If
            
            ElseIf (VarType(txtconteudo.Text)) = vbString Then
                txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & "'" & txtconteudo & "'"
            
            End If
            
         
           
         'Select Case
      '   If DataGrid1.Text  vbInteger Then
         'Case vbInteger
              '  txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & "'" & txtconteudo & "'"
          'End If
'          Case vbString
'          txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'
          
          
'            Case txtCondicaoUsu <> ""
 
                'Case dbBigInt
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case dbBinary
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case dbBoolean
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case dbByte
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case dbChar
'                    txtCondicaoUsu.Text = lstCampo.List(i) & "" & cmbCondicao.Text & "'" & txtConteudo & "'"
'                Case dbCurrency
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case dbDate
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & "#" & txtConteudo & "#"
'                Case dbDecimal
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case dbDouble
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case dbFloat
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case dbGUID
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case dbInteger
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case dbLong
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case dbLongBinary
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case dbMemo
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & "'" & txtConteudo & "'"
'                Case dbNumeric
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case dbSingle
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case dbText
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & "'" & txtConteudo & "'"
'                Case dbTime
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & "#" & txtConteudo & "#"
'                Case dbTimeStamp
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & "#" & txtConteudo & "#"
'                Case dbVarBinary
'                    txtCondicaoUsu.Text = lstCampo.List(i) & " " & cmbCondicao.Text & " " & txtConteudo
'                Case Else
'                    MsgBox "Tipo Inválido !", 48, Me.Caption
'                    Exit Sub
       '  End Select
            Exit For
        End If
    Next i
    
    FMontaChave

End Sub

Private Sub DataGrid1_DblClick()
On Error GoTo TratamentoDeErro


    Dim wColuna As Double, i
    wColuna = (lstRetorna.List(0))
    DataGrid1.Col = wColuna
    wBuscaRetorna = DataGrid1.Text
    For i = 0 To IIf((lstTitulo.ListCount - 1) > 10, 10, lstTitulo.ListCount - 1)
        DataGrid1.Col = i
        WxRetorno(i) = DataGrid1.Text
    Next i
    Unload Me

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



Private Sub Form_Load()
    Me.Caption = "ATB Info - " + Me.Caption
    wBuscaRetorna = Empty
    For i = 0 To 5
       WxRetorno(i) = Empty
    Next
End Sub



Private Sub tobLocalizar_ButtonClick(ByVal Button As ComctlLib.Button)

    Dim sOrdem As String
    Dim i As Integer
    
    On Error GoTo Sair
    
    Select Case Format(Button.Key, ">")
        
        Case "FECHAR"
            For i = 0 To 5
               WxRetorno(i) = Empty
            Next i
            Unload Me
            Exit Sub
            
        Case "ORDENARAZ"
            FOrdenaColunas "ASC"
        
        Case "ORDENARZA"
            FOrdenaColunas "DESC"
        
        Case "FILTRAR"
            FPosFiltrar
            FCarregaOpcoesFiltro
        
        Case "LOCALIZA"
            FPosLocalizar
            txtPalavra.SetFocus
        
'        Case "PRIMEIRO"
'            DataGrid1.reco.MoveFirst
'
'        Case "ANTERIOR"
'            If DataGrid1.Recordset.BOF = True Then
'                MSHFlexGrid1.Recordset.MoveLast
'            Else
'                MSHFlexGrid1.Recordset.MovePrevious
'            End If
'
'        Case "PROXIMO"
'            If dtaDados.Recordset.EOF = True Then
'                dtaDados.Recordset.MoveFirst
'            Else
'                dtaDados.Recordset.MoveNext
'            End If
        
'        Case "ULTIMO"
'            dtaDados.Recordset.MoveLast
    
    End Select
    
    Exit Sub

Sair:
    
    MsgBox "Não foi possivel executar esta opção no momento !" _
            & Chr$(13) & "Tente novamente.", 48, Me.Caption
    
    Exit Sub

End Sub

Private Sub txtConteudo_Change()
    cmdFiltra.Enabled = True
    
    
End Sub



Private Sub txtconteudo_LostFocus()
    txtconteudo = Replace(txtconteudo, ",", ".")
End Sub

Private Sub txtPalavra_Change()

'    Dim i As Integer
'    Dim iCol As Integer
'    Dim iRow As Integer
'    Dim sCampo As String
'    Dim sCriterio As String
'
'    If dbgDados.SelStartCol = -1 Or dbgDados.SelEndCol = -1 Then
'        MsgBox "Selecione uma coluna !", 48, "Aviso"
'        Exit Sub
'    End If
'
'    If dbgDados.SelStartCol <> dbgDados.SelEndCol = -1 Then
'        MsgBox "Selecione apenas uma coluna !", 48, "Aviso"
'        Exit Sub
'    End If
'
'    If txtPalavra.Text = Empty Then Exit Sub
'
'    With dbgDados
'
'        iCol = .Col
'        iRow = .Row
'
'        .Col = dbgDados.SelStartCol
'
'        sCampo = frmLocalizar.lstTitulo.List(dbgDados.SelStartCol)
'
'        If .Col <= -1 Then
'            MsgBox "Selecione uma coluna para procura !", 48, Me.Caption
'            Exit Sub
'        End If
'
'        Select Case dtaDados.Recordset.Fields(.Col).Type
'            Case dbBigInt
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbBinary
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbBoolean
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbByte
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbChar
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") ='" & Format(txtPalavra.Text, ">") & "'"
'            Case dbCurrency
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbDate
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") = #" & Format(txtPalavra.Text, ">") & "#"
'            Case dbDecimal
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbDouble
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbFloat
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbGUID
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbInteger
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbLong
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbLongBinary
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbMemo
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbNumeric
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbSingle
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case dbText
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") ='" & Format(txtPalavra.Text, ">") & "'"
'            Case dbTime
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =#" & Format(txtPalavra.Text, ">") & "#"
'            Case dbTimeStamp
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =#" & Format(txtPalavra.Text, ">") & "#"
'            Case dbVarBinary
'                sCriterio = "Mid(" & sCampo & ", 1, " & Len(txtPalavra.Text) & ") =" & Format(txtPalavra.Text, ">")
'            Case Else
'                MsgBox "Tipo Inválido !", 48, Me.Caption
'                Exit Sub
'        End Select
'
'        If dtaDados.Recordset.Fields(.Col).Type = dbDate Then
'            If Len(txtPalavra.Text) < 8 Or Len(txtPalavra.Text) < 10 Then
'            Else
'                dtaDados.Recordset.FindFirst sCriterio
'            End If
'        Else
'            On Error Resume Next
'            dtaDados.Recordset.FindFirst sCriterio
'        End If
'
'    End With
'
'    txtPalavra.SetFocus

End Sub

Private Sub txtPalavra_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    dbgDados_DblClick
'End If
End Sub


Private Function PaginarRecordset(recset As ADODB.Recordset) As ADODB.Recordset
  Dim subRst As New ADODB.Recordset
  Dim x As Long
  Dim fld As Field
  Dim origPage As Long
   Dim rst As ADODB.Recordset
  'Set subRst = New ADODB.Recordset
  Set rst = New ADODB.Recordset
  Const PAGE_SIZE = 200
  
  origPage = IIf(recset.AbsolutePage > 0, recset.AbsolutePage, 1)
  
  With subRst
    If .State = adStateOpen Then .Close
    'Cria campos
    For Each fld In recset.Fields
      .Fields.Append fld.Name, fld.Type, fld.DefinedSize, fld.Attributes
    Next fld
    'Inclui registros
    .Open
    For x = 1 To PAGE_SIZE
      If recset.EOF Then Exit For
      .AddNew
      For Each fld In recset.Fields
        subRst(fld.Name) = fld.Value
      Next fld
      .Update
      recset.MoveNext
    Next x
    .MoveFirst
    recset.AbsolutePage = origPage
  End With
  Set PaginarRecordset = subRst

End Function

