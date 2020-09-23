VERSION 5.00
Begin VB.Form SForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   Picture         =   "SForm.frx":0000
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   252
      Left            =   2160
      TabIndex        =   0
      Top             =   2040
      Width           =   1212
   End
End
Attribute VB_Name = "SForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
Hide
OForm.Visible = True
Unload Me
End Sub
