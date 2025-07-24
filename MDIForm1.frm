VERSION 5.00
Begin VB.MDIForm frm_Principal 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "XYZ - Administradora de Cartões de Crédito  - Gerenciamento de Transações"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   10080
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "MDIForm1.frx":1084A
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


