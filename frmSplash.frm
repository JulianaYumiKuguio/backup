VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3795
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   14160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":1084A
   ScaleHeight     =   3795
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerSplash 
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "XYZ - Administradora de Cartões - Gerenciamento de Transações (Crédito)"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   8040
      TabIndex        =   2
      Top             =   720
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Candidata:  Juliana Yumi Kuguio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   615
      Left            =   9240
      TabIndex        =   1
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Label Lbl_Aviso 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conectando ao Sistema ...  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   9840
      TabIndex        =   0
      Top             =   3120
      Width           =   4125
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Activate()
    '' Abre a conexão Através da configuração do arquivo config
    '' Onde é informado os parâmetros (Servidor, Usuário, Senha, Banco de Dados, Caminho do RPT Crystal, Caminho do Log de Erros)
    '' O arquivo deve ficar na mesma pasta do projeto
    '' Criei módulo de conexão Module1
    Call AbreConexaoBD
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub



Private Sub TimerSplash_Timer()

    TimerSplash.Enabled = False
    
    Unload Me
    Call fChamarTelaPrincipal


End Sub

Public Function fChamarTelaPrincipal()
    frm_Principal.Show
End Function

