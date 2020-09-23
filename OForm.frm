VERSION 5.00
Begin VB.Form OForm 
   Caption         =   "Options"
   ClientHeight    =   4548
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   6012
   Icon            =   "OForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4548
   ScaleWidth      =   6012
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Controls :"
      Height          =   1452
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   4332
      Begin VB.Label Label3 
         Caption         =   "Left/Right controls turning"
         Height          =   252
         Index           =   4
         Left            =   2040
         TabIndex        =   11
         Top             =   840
         Width           =   2172
      End
      Begin VB.Label Label3 
         Caption         =   "Up/Down controls pitch"
         Height          =   252
         Index           =   3
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   1932
      End
      Begin VB.Label Label3 
         Caption         =   "M = Roll Right"
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1332
      End
      Begin VB.Label Label3 
         Caption         =   "N = Roll Left"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label Label3 
         Caption         =   "Z = Slow Down"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1932
      End
      Begin VB.Label Label2 
         Caption         =   "A = Speed Up"
         Height          =   252
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1932
      End
   End
   Begin VB.Frame ExtrasF 
      Caption         =   "Extra Features"
      Height          =   972
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   5772
      Begin VB.OptionButton UFOO 
         Caption         =   "Stars, Comets and UFO's (Slowest)"
         Height          =   492
         Left            =   3720
         TabIndex        =   14
         Top             =   240
         Width           =   1692
      End
      Begin VB.OptionButton CometsO 
         Caption         =   "Stars and Comets (Fast)"
         Height          =   492
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   1932
      End
      Begin VB.OptionButton StarsO 
         Caption         =   "Stars Only (Fastest)"
         Height          =   252
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1692
      End
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "GO !!!"
      Height          =   492
      Left            =   4560
      TabIndex        =   3
      ToolTipText     =   "Sets you off on your galactic journey"
      Top             =   3600
      Width           =   1332
   End
   Begin VB.Frame DensityF 
      Caption         =   "Galaxy Density = 500"
      Height          =   1692
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5772
      Begin VB.HScrollBar Density 
         Height          =   492
         LargeChange     =   100
         Left            =   120
         Max             =   5000
         Min             =   10
         SmallChange     =   10
         TabIndex        =   1
         Top             =   360
         Value           =   500
         Width           =   5532
      End
      Begin VB.Label Label1 
         Caption         =   $"OForm.frx":030A
         Height          =   612
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   5532
      End
   End
End
Attribute VB_Name = "OForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGO_Click()
NUMSTARS = Density.Value
If StarsO Then
Mode = STARS
  Else
  If CometsO Then
  Mode = COMETS
          Else
          Mode = UFOS
        End If
End If
Hide
MsgBox "When you want to end, press Escape to exit the program", vbInformation, "How to Exit!"
GForm.Visible = True
Unload Me
End Sub

Private Sub Density_Change()
DensityF.Caption = "Density = " & Density.Value
End Sub
