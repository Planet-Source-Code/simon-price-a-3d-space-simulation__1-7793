VERSION 5.00
Begin VB.Form GForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5556
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6276
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   523
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox UFOmask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   1200
      Picture         =   "GForm.frx":0000
      ScaleHeight     =   372
      ScaleWidth      =   612
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.PictureBox UFOpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   360
      Picture         =   "GForm.frx":0A5A
      ScaleHeight     =   372
      ScaleWidth      =   612
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.PictureBox PB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   3732
      Left            =   3000
      ScaleHeight     =   311
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   431
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   5172
   End
End
Attribute VB_Name = "GForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeySpace
  
  Case vbKeyC
  PB.Picture = PB.Image
  SavePicture PB.Picture, App.Path & "/spacepic.bmp"
  PB.Picture = LoadPicture()

  Case vbKeyEscape
  MsgBox "Thankyou for playing space by Simon Price. This space simulation took me 3 hours to make. To see games which I have spend more time on, visit my website : www.VBgames.co.uk", vbInformation, "Thankyou for playing"
    End
  
  Case vbKeyA
  Ship.zm = Ship.zm + 1
  Case vbKeyZ
  If Ship.zm > 2 Then Ship.zm = Ship.zm - 1
  
  Case vbKeyN
        Select Case Ship.RollAngle
          Case 0
          Ship.RollAngle = 359
        Case 0 To 10
          Ship.RollAngle = Ship.RollAngle - 1
        Case 351 To 359
          Ship.RollAngle = Ship.RollAngle - 1
        End Select
  
  Case vbKeyM
        Select Case Ship.RollAngle
          Case 359
          Ship.RollAngle = 0
        Case 0 To 9
          Ship.RollAngle = Ship.RollAngle + 1
        Case 350 To 359
          Ship.RollAngle = Ship.RollAngle + 1
        End Select
  
  Case vbKeyUp
        Select Case Ship.PitchAngle
          Case 359
          Ship.PitchAngle = 0
        Case 0
          Ship.PitchAngle = Ship.PitchAngle + 1
        Case 350 To 351
          Ship.PitchAngle = Ship.PitchAngle + 1
        End Select
  
  Case vbKeyDown
        Select Case Ship.PitchAngle
          Case 0
          Ship.PitchAngle = 359
        Case 0 To 1
          Ship.PitchAngle = Ship.PitchAngle - 1
        Case 350
          Ship.PitchAngle = Ship.PitchAngle - 1
        End Select
       
  Case vbKeyLeft
        Select Case Ship.TurnAngle
          Case 359
          Ship.TurnAngle = 0
        Case 0
          Ship.TurnAngle = Ship.TurnAngle + 1
        Case 350 To 351
          Ship.TurnAngle = Ship.TurnAngle + 1
        End Select
  
  Case vbKeyRight
        Select Case Ship.TurnAngle
          Case 0
          Ship.TurnAngle = 359
        Case 0 To 1
          Ship.TurnAngle = Ship.TurnAngle - 1
        Case 350
          Ship.TurnAngle = Ship.TurnAngle - 1
        End Select
       
End Select
End Sub

Private Sub Form_Load()
SortLayout
Show
BuildTrigTable
CreateStars
Ship.zm = 1
Select Case Mode
  Case STARS
    DoStars
  Case COMETS
    DoStarsAndComets
  Case UFOS
    DoStarsAndCometsAndUFOs
End Select
End Sub

Private Sub SortLayout()
Move 0, 0, Screen.Width, Screen.Height
PB.Move 0, 0, 600, 450
DispWidth = Screen.Height * 1.36 / Screen.TwipsPerPixelY
DispHeight = Screen.Height * 0.977 / Screen.TwipsPerPixelY
End Sub

Private Sub CreateStars()
ReDim Star(1 To NUMSTARS)
For i = 1 To NUMSTARS
  Star(i).X = Rnd * VIEWWIDTH - VIEWWIDTH \ 2
  Star(i).Y = Rnd * VIEWHEIGHT - VIEWHEIGHT \ 2
  Star(i).z = Rnd * VIEWDEPTH
Next
End Sub

Private Sub DoStars()
On Error Resume Next
Do
DoEvents
PB.Cls

For i = 1 To NUMSTARS
'move them
Y = Star(i).Y * Cosine(Ship.RollAngle) - Star(i).X * Sine(Ship.RollAngle)
Star(i).X = Star(i).X * Cosine(Ship.RollAngle) + Star(i).Y * Sine(Ship.RollAngle)
Star(i).Y = Y
  
z = Star(i).z * Cosine(Ship.PitchAngle) - Star(i).Y * Sine(Ship.PitchAngle)
Star(i).Y = Star(i).Y * Cosine(Ship.PitchAngle) + Star(i).z * Sine(Ship.PitchAngle)
Star(i).z = z
  
Star(i).X = Star(i).X * Cosine(Ship.TurnAngle) - Star(i).z * Sine(Ship.TurnAngle)
Star(i).z = Star(i).z * Cosine(Ship.TurnAngle) + Star(i).X * Sine(Ship.TurnAngle)
  
  If Star(i).z <= 0 Then
  Star(i).z = VIEWDEPTH
  Star(i).X = Rnd * VIEWWIDTH - VIEWWIDTH \ 2
  Star(i).Y = Rnd * VIEWHEIGHT - VIEWHEIGHT \ 2
  Else
  LensDivDist = LENS / Star(i).z
  X = CX + Star(i).X * LensDivDist
  Y = CY - Star(i).Y * LensDivDist
  Select Case X
    Case 0 To PBWIDTH
      Select Case Y
        Case 0 To PBHEIGHT
          MoveToEx PB.hdc, X, Y, lpPoint
        Case Else
          Star(i).z = Star(i).z - Ship.zm
          GoTo NextOne
      End Select
    Case Else
      Star(i).z = Star(i).z - Ship.zm
      GoTo NextOne
  End Select
  Star(i).z = Star(i).z - Ship.zm
'draw them
  LensDivDist = LENS / Star(i).z
  X = CX + Star(i).X * LensDivDist
  Y = CY - Star(i).Y * LensDivDist
  Select Case X
    Case 0 To PBWIDTH
      Select Case Y
        Case 0 To PBHEIGHT
          LineTo PB.hdc, X, Y
      End Select
  End Select
  End If
NextOne:
Next
StretchBlt hdc, 0, 0, DispWidth, DispHeight, PB.hdc, 0, 0, PBWIDTH, PBHEIGHT, vbSrcCopy
Loop
End Sub

Private Sub DoStarsAndCometsAndUFOs()
On Error Resume Next
Do
DoEvents
PB.Cls

For i = 1 To NUMSTARS
'move them
Y = Star(i).Y * Cosine(Ship.RollAngle) - Star(i).X * Sine(Ship.RollAngle)
Star(i).X = Star(i).X * Cosine(Ship.RollAngle) + Star(i).Y * Sine(Ship.RollAngle)
Star(i).Y = Y
  
z = Star(i).z * Cosine(Ship.PitchAngle) - Star(i).Y * Sine(Ship.PitchAngle)
Star(i).Y = Star(i).Y * Cosine(Ship.PitchAngle) + Star(i).z * Sine(Ship.PitchAngle)
Star(i).z = z
  
Star(i).X = Star(i).X * Cosine(Ship.TurnAngle) - Star(i).z * Sine(Ship.TurnAngle)
Star(i).z = Star(i).z * Cosine(Ship.TurnAngle) + Star(i).X * Sine(Ship.TurnAngle)
  If Star(i).z <= 0 Then
  Star(i).z = VIEWDEPTH
  Star(i).X = Rnd * VIEWWIDTH - VIEWWIDTH \ 2
  Star(i).Y = Rnd * VIEWHEIGHT - VIEWHEIGHT \ 2
  Else
  LensDivDist = LENS / Star(i).z
  X = CX + Star(i).X * LensDivDist
  Y = CY - Star(i).Y * LensDivDist
  Select Case X
    Case 0 To PBWIDTH
      Select Case Y
        Case 0 To PBHEIGHT
          MoveToEx PB.hdc, X, Y, lpPoint
        Case Else
          Star(i).z = Star(i).z - Ship.zm
          GoTo NextOne
      End Select
    Case Else
      Star(i).z = Star(i).z - Ship.zm
      GoTo NextOne
  End Select
  Star(i).z = Star(i).z - Ship.zm
'draw them
  LensDivDist = LENS / Star(i).z
  X = CX + Star(i).X * LensDivDist
  Y = CY - Star(i).Y * LensDivDist
  z = Star(i).z
  Select Case X
    Case 0 To PBWIDTH
      Select Case Y
        Case 0 To PBHEIGHT
          Select Case i Mod 100
          Case 0
'            PB.DrawWidth = 2
'            PB.ForeColor = vbYellow
'            LineTo PB.hdc, X, Y
'            PB.ForeColor = vbWhite
'            PB.DrawWidth = 1
          w = (1000 - X - Y - z) \ 3 * 6
          h = (1000 - X - Y - z) \ 3 * 10
          X = X - w \ 2
          Y = Y - h \ 2
             StretchBlt PB.hdc, X, Y, w, h, UFOmask.hdc, 0, 0, 50, 30, vbMergePaint
             StretchBlt PB.hdc, X, Y, w, h, UFOpic.hdc, 0, 0, 50, 30, vbSrcAnd
          Case 50
            PB.DrawWidth = 2
            PB.ForeColor = vbCyan
            LineTo PB.hdc, X, Y
            PB.ForeColor = vbWhite
            PB.DrawWidth = 1
'          w = (1000 - X - Y - z) \ 3 * 6
'          h = (1000 - X - Y - z) \ 3 * 10
'          X = X - w \ 2
'          Y = Y - h \ 2
'             StretchBlt PB.hdc, X, Y, w, h, UFOmask.hdc, 0, 0, 50, 30, vbMergePaint
'             StretchBlt PB.hdc, X, Y, w, h, UFOpic.hdc, 0, 0, 50, 30, vbSrcAnd
          Case Else
            LineTo PB.hdc, X, Y
          End Select
      End Select
  End Select
  End If
NextOne:
Next
StretchBlt hdc, 0, 0, DispWidth, DispHeight, PB.hdc, 0, 0, PBWIDTH, PBHEIGHT, vbSrcCopy
Loop
End Sub

Private Sub DoStarsAndComets()
On Error Resume Next
Do
DoEvents
PB.Cls

For i = 1 To NUMSTARS
'move them
Y = Star(i).Y * Cosine(Ship.RollAngle) - Star(i).X * Sine(Ship.RollAngle)
Star(i).X = Star(i).X * Cosine(Ship.RollAngle) + Star(i).Y * Sine(Ship.RollAngle)
Star(i).Y = Y
  
z = Star(i).z * Cosine(Ship.PitchAngle) - Star(i).Y * Sine(Ship.PitchAngle)
Star(i).Y = Star(i).Y * Cosine(Ship.PitchAngle) + Star(i).z * Sine(Ship.PitchAngle)
Star(i).z = z
  
Star(i).X = Star(i).X * Cosine(Ship.TurnAngle) - Star(i).z * Sine(Ship.TurnAngle)
Star(i).z = Star(i).z * Cosine(Ship.TurnAngle) + Star(i).X * Sine(Ship.TurnAngle)
  If Star(i).z <= 0 Then
  Star(i).z = VIEWDEPTH
  Star(i).X = Rnd * VIEWWIDTH - VIEWWIDTH \ 2
  Star(i).Y = Rnd * VIEWHEIGHT - VIEWHEIGHT \ 2
  Else
  LensDivDist = LENS / Star(i).z
  X = CX + Star(i).X * LensDivDist
  Y = CY - Star(i).Y * LensDivDist
  Select Case X
    Case 0 To PBWIDTH
      Select Case Y
        Case 0 To PBHEIGHT
          MoveToEx PB.hdc, X, Y, lpPoint
        Case Else
          Star(i).z = Star(i).z - Ship.zm
          GoTo NextOne
      End Select
    Case Else
      Star(i).z = Star(i).z - Ship.zm
      GoTo NextOne
  End Select
  Star(i).z = Star(i).z - Ship.zm
'draw them
  LensDivDist = LENS / Star(i).z
  X = CX + Star(i).X * LensDivDist
  Y = CY - Star(i).Y * LensDivDist
  z = Star(i).z
  Select Case X
    Case 0 To PBWIDTH
      Select Case Y
        Case 0 To PBHEIGHT
          Select Case i Mod 100
          Case 0
            PB.DrawWidth = 2
            PB.ForeColor = vbYellow
            LineTo PB.hdc, X, Y
            PB.ForeColor = vbWhite
            PB.DrawWidth = 1
             StretchBlt PB.hdc, X, Y, w, h, UFOmask.hdc, 0, 0, 50, 30, vbMergePaint
             StretchBlt PB.hdc, X, Y, w, h, UFOpic.hdc, 0, 0, 50, 30, vbSrcAnd
          Case 50
            PB.DrawWidth = 2
            PB.ForeColor = vbCyan
            LineTo PB.hdc, X, Y
            PB.ForeColor = vbWhite
            PB.DrawWidth = 1
          Case Else
            LineTo PB.hdc, X, Y
          End Select
      End Select
  End Select
  End If
NextOne:
Next
StretchBlt hdc, 0, 0, DispWidth, DispHeight, PB.hdc, 0, 0, PBWIDTH, PBHEIGHT, vbSrcCopy
Loop
End Sub

