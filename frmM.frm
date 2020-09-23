VERSION 5.00
Begin VB.Form frmM 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   734
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chSCIA 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TRAIL"
      Height          =   375
      Left            =   14040
      TabIndex        =   62
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CheckBox chCAM 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Follow"
      Height          =   375
      Left            =   14040
      TabIndex        =   61
      Top             =   6600
      Width           =   1095
   End
   Begin VB.HScrollBar scrollH 
      Height          =   255
      Left            =   14040
      Max             =   12
      TabIndex        =   60
      Top             =   7200
      Width           =   975
   End
   Begin VB.HScrollBar ScrollTEAM 
      Height          =   255
      Left            =   14040
      Max             =   2
      Min             =   1
      TabIndex        =   59
      Top             =   6960
      Value           =   1
      Width           =   975
   End
   Begin VB.CheckBox chFIXEnemy 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fix Enemy"
      Height          =   375
      Left            =   14040
      TabIndex        =   51
      ToolTipText     =   "Do not Change Enemy until it has been hit by you or Friends. Then Choose nearest Enemy as Fix Enemy."
      Top             =   6000
      Width           =   1095
   End
   Begin VB.PictureBox pPOP2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   14040
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   47
      ToolTipText     =   "Eyes Inputs"
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox pPop1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   14040
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   46
      ToolTipText     =   "Eyes Inputs"
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox PicN 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   11520
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   44
      Top             =   8520
      Width           =   3495
   End
   Begin VB.PictureBox PicRes 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   11520
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   159
      TabIndex        =   43
      Top             =   6600
      Width           =   2415
   End
   Begin VB.PictureBox PicWin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11520
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   42
      ToolTipText     =   "Eyes Inputs"
      Top             =   6360
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0C000&
      Height          =   285
      Left            =   12720
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Gen AVG"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H0000FF00&
      Height          =   285
      Left            =   12720
      TabIndex        =   20
      Text            =   "Gen Bestift"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   12720
      TabIndex        =   19
      Text            =   "NEWrandom"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   12720
      TabIndex        =   18
      Text            =   "mut"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   12720
      TabIndex        =   17
      Text            =   "Acc"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   12720
      TabIndex        =   16
      Text            =   "G"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.PictureBox PicOUT 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   753
      TabIndex        =   15
      ToolTipText     =   "Outputs"
      Top             =   9960
      Width           =   11295
   End
   Begin VB.TextBox GEN 
      Height          =   285
      Left            =   12720
      TabIndex        =   8
      Text            =   "G"
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox ACC 
      Height          =   285
      Left            =   12720
      TabIndex        =   7
      Text            =   "Acc"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox MUT 
      Height          =   285
      Left            =   12720
      TabIndex        =   6
      Text            =   "mut"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox NEWr 
      Height          =   285
      Left            =   12720
      TabIndex        =   5
      Text            =   "NEWrandom"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox BFIT 
      BackColor       =   &H0000FF00&
      Height          =   285
      Left            =   12720
      TabIndex        =   4
      Text            =   "Gen Bestift"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox gAVG 
      BackColor       =   &H00C0C000&
      Height          =   285
      Left            =   12720
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Gen AVG"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox PicInput 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   753
      TabIndex        =   2
      ToolTipText     =   "Eyes Inputs"
      Top             =   9480
      Width           =   11295
   End
   Begin VB.Timer TimerMain 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10440
      Top             =   480
   End
   Begin VB.TextBox INFO 
      Height          =   495
      Left            =   11640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmM.frx":0000
      Top             =   120
      Width           =   3375
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   9015
      Left            =   120
      ScaleHeight     =   601
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   753
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "^"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   8640
         TabIndex        =   58
         Top             =   7920
         Width           =   495
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   8640
         TabIndex        =   57
         Top             =   8520
         Width           =   495
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   8040
         TabIndex        =   56
         Top             =   8160
         Width           =   495
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   9240
         TabIndex        =   55
         Top             =   8160
         Width           =   495
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   9840
         TabIndex        =   54
         Top             =   8160
         Width           =   495
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   10560
         TabIndex        =   53
         Top             =   8160
         Width           =   495
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   10320
         TabIndex        =   52
         Top             =   8280
         Width           =   255
      End
   End
   Begin VB.Label InpDesc 
      Alignment       =   2  'Center
      Caption         =   "DIR-DIR"
      Height          =   255
      Index           =   10
      Left            =   9840
      TabIndex        =   64
      ToolTipText     =   "Enemy Direction (to mine)"
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Inputs (red), Outputs (green)  and  Neural Network Activity of the best of Red Team"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   63
      Top             =   10320
      Width           =   4335
   End
   Begin VB.Label InpDesc 
      Alignment       =   2  'Center
      Caption         =   "E VEL"
      Height          =   255
      Index           =   9
      Left            =   8640
      TabIndex        =   50
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label InpDesc 
      Alignment       =   2  'Center
      Caption         =   "my VEL"
      Height          =   255
      Index           =   8
      Left            =   7680
      TabIndex        =   49
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label InpDesc 
      Alignment       =   2  'Center
      Caption         =   "OR-OR"
      Height          =   255
      Index           =   7
      Left            =   6240
      TabIndex        =   48
      ToolTipText     =   "Enemy Orientation (to me)"
      Top             =   9240
      Width           =   1455
   End
   Begin VB.Label Lprewin 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   11520
      TabIndex        =   45
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label lAVG2 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   12840
      TabIndex        =   41
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lAVG1 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   11640
      TabIndex        =   40
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label OutDesc 
      Alignment       =   2  'Center
      Caption         =   "Fire Shot"
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   39
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Label OutDesc 
      Alignment       =   2  'Center
      Caption         =   "Right Rocket"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   38
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Label OutDesc 
      Alignment       =   2  'Center
      Caption         =   "Left Rocket"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Label InpDesc 
      Alignment       =   2  'Center
      Caption         =   "Av. Shots"
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   36
      ToolTipText     =   "Available Shots"
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label InpDesc 
      Alignment       =   2  'Center
      Caption         =   "(E) Shot Dist"
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   35
      Top             =   9240
      Width           =   1095
   End
   Begin VB.Label InpDesc 
      Alignment       =   2  'Center
      Caption         =   "(E) Shot R"
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   34
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label InpDesc 
      Alignment       =   2  'Center
      Caption         =   "(E) Shot L"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   33
      Top             =   9240
      Width           =   855
   End
   Begin VB.Label InpDesc 
      Alignment       =   2  'Center
      Caption         =   "E DIST"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   32
      Top             =   9240
      Width           =   855
   End
   Begin VB.Label InpDesc 
      Alignment       =   2  'Center
      Caption         =   "E Right"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   31
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label InpDesc 
      Alignment       =   2  'Center
      Caption         =   "E Left"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   9240
      Width           =   735
   End
   Begin VB.Label Lwin2 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   12840
      TabIndex        =   29
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Lwin1 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   11640
      TabIndex        =   28
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "G Best Fit"
      Height          =   255
      Left            =   11640
      TabIndex        =   27
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "New Random"
      Height          =   255
      Left            =   11640
      TabIndex        =   26
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mutations"
      Height          =   255
      Left            =   11640
      TabIndex        =   25
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reproductions"
      Height          =   255
      Left            =   11640
      TabIndex        =   24
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "G Avg Fit"
      Height          =   255
      Left            =   11640
      TabIndex        =   23
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generation"
      Height          =   255
      Left            =   11640
      TabIndex        =   22
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generation"
      Height          =   255
      Left            =   11640
      TabIndex        =   14
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "G Avg Fit"
      Height          =   255
      Left            =   11640
      TabIndex        =   13
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reproductions"
      Height          =   255
      Left            =   11640
      TabIndex        =   12
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mutations"
      Height          =   255
      Left            =   11640
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "New Random"
      Height          =   255
      Left            =   11640
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      Caption         =   "G Best Fit"
      Height          =   255
      Left            =   11640
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
End
Attribute VB_Name = "frmM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Option Base 1


Private BR1            As New simplyBrainsPOP
Private BR2            As New simplyBrainsPOP

Private I              As Long


Public IndexBest1      As Long
Public IndexBest2      As Long

Private Win1           As Long
Private Win2           As Long

Private AVG1           As Single
Private AVG2           As Single

Private KpicWin        As Single

Private PY             As Single
Private oX             As Integer
Private oY             As Integer



Private Sub chSCIA_Click()
If chSCIA Then DoScia = True Else: DoScia = False

End Sub

Private Sub cmdNAVIGATE_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Navigating = True

    Do

        Select Case Index
            Case 0
                PanZoomChanged = True
                PanY = PanY - 0.001 / ZOOM

            Case 1
                PanZoomChanged = True
                PanY = PanY + 0.001 / ZOOM


            Case 2
                PanZoomChanged = True
                PanX = PanX - 0.001 / ZOOM

            Case 3
                PanZoomChanged = True
                PanX = PanX + 0.001 / ZOOM


            Case 4
                PanZoomChanged = True
                ZOOM = ZOOM / 1.000002
            Case 5
                PanZoomChanged = True
                ZOOM = ZOOM * 1.000002    '1.2

            Case 6
                PanZoomChanged = True
                ZOOM = 1
                PanX = MaxX / 2
                PanY = MaxY / 2
                ZOOM = PIC.Height / MaxY
        End Select

        '        SW.DRAW
        '        PIC.Refresh

        DoEvents
InvZoom = 1 / ZOOM
    Loop While Navigating


End Sub

Private Sub cmdNAVIGATE_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Navigating = False
End Sub

Private Sub Form_Activate()

    DoEvents

    MAINLOOP


End Sub

Private Sub Form_Load()



    MaxXPic = PIC.Width
    MaxYpic = PIC.Height
    CenX = MaxXPic \ 2
    CenY = MaxYpic \ 2

    MaxX = PIC.Width * 2
    MaxY = PIC.Height * 2

    PanX = MaxX \ 2               'CenX
    PanY = MaxY \ 2               'CenY
    ZOOM = PIC.Height / MaxY
    InvZoom = 1 / ZOOM

    'ZOOM = 0.45             '0.5 '0.4 ' 0.55    '0.65    '0.5 '0.75
    'MaxX = PIC.Width / ZOOM
    'MaxY = PIC.Height / ZOOM

    scrollH.Max = NH


    KpicWin = (PicWin.Width \ 2) / MaxFitDist


    For I = 0 To NofInputs - 1
        InpDesc(I).Width = PicInput.Width / NofInputs
        InpDesc(I).Left = PicInput.Left + I * PicInput.Width / NofInputs
    Next
    For I = 0 To 3 - 1
        OutDesc(I).Width = PicOUT.Width / 3
        OutDesc(I).Left = PicOUT.Left + I * PicInput.Width / 3
    Next



    Randomize Timer

    IndexBest1 = 1
    IndexBest2 = 1




    Randomize Timer


    ReDim H1(NH)
    ReDim H2(NH)

    For I = 1 To NH
        With H1(I)
            .SurName = UCase(GenName(3, 4))
            .Name = GenName(3, 6)
            .Gender = IIf(Rnd < 0.5, True, False)
            .PosX = Rnd * MaxX
            .PosY = Rnd * MaxY
            .ANG = Rnd * PI2
            .ACCLeft = 0.1        '4
            .ACCRight = 0.1
            .R = 10               ' 8           '7
            .AvailShots = MaxShots
            .EnemyToAttack = IndexBest2
            '.SciaLen = 29
        End With
        With H2(I)
            .SurName = UCase(GenName(3, 4))
            .Name = GenName(3, 6)
            .Gender = IIf(Rnd < 0.5, True, False)
            .PosX = Rnd * MaxX
            .PosY = Rnd * MaxY
            .ANG = Rnd * PI2
            .ACCLeft = 0.1        '4
            .ACCRight = 0.1
            .R = 10               '8           '7
            .AvailShots = MaxShots
            .EnemyToAttack = IndexBest1
            '.SciaLen = 29
        End With
    Next


    NF = 10                       ' '10 '14 '12 '8
    ReDim F(NF)

    For I = 1 To NF
        With F(I)
            .PosX = Rnd * MaxX
            .PosY = Rnd * MaxY
            .R = 10 + Rnd * 20
        End With
    Next



    'Br1= Polulation of brains for team1
    'Br2= Polulation of brains for team2

    'Each Brain have 3 "cells"
    BR1.InitBrains NH, 3
    BR2.InitBrains NH, 3

    'First one have only 2 output
    BR1.InitBrainCell 1, Array(NofInputs, 2), 1, PicN    '5
    BR2.InitBrainCell 1, Array(NofInputs, 2), 1, PicN    '7, 4), 12

    'The outputs of first Cell
    'determinate if run the 2nd or 3rd cell to
    'determinate to output movement
    'SEARCH:If Out1(1) > Out1(2) Then
    '       H1(I).BRtoRUN = 2
    '       Else
    '       H1(I).BRtoRUN = 3
    '       End If
    BR1.InitBrainCell 2, Array(NofInputs, 6, 3), 5, PicN    '5
    BR2.InitBrainCell 2, Array(NofInputs, 6, 3), 5, PicN    '7, 4), 12
    BR1.InitBrainCell 3, Array(NofInputs, 6, 3), 5, PicN    '5
    BR2.InitBrainCell 3, Array(NofInputs, 6, 3), 5, PicN    '7, 4), 12


    '    GA1.INIT1_EvolutionParams MutateAll, 0.05, 0.25, 0.25, SelWheel, CrossBySwap, SonToWorst, False, 10
    '    GA1.INIT2_Pop NH, 0, 100, BR1.GetNofTotalGenes, BR1.GetNofTotalGenes    ', True
    '    GA2.INIT1_EvolutionParams MutateAll, 0.05, 0.25, 0.25, SelWheel, CrossBySwap, SonToWorst, False, 10
    '    GA2.INIT2_Pop NH, 0, 100, BR2.GetNofTotalGenes, BR1.GetNofTotalGenes    ', True


    GA1.INIT1_EvolutionParams MutateAll, 0.05, 0.1, 0.2, 0.25, SelWheel, CrossBySwap, SonToWorst, False, 5, TDbyIdenticalIndi
    GA1.INIT2_Pop NH, 0, 100, BR1.GetNofTotalGenes, BR1.GetNofTotalGenes    ', True
    GA2.INIT1_EvolutionParams MutateAll, 0.05, 0.1, 0.2, 0.25, SelWheel, CrossBySwap, SonToWorst, False, 5, TDbyIdenticalIndi
    GA2.INIT2_Pop NH, 0, 100, BR2.GetNofTotalGenes, BR1.GetNofTotalGenes    ', True

    '************

    If Dir(App.Path & "\pop1.txt") <> "" Then GA1.LoadPOP "pop1.txt"
    If Dir(App.Path & "\pop2.txt") <> "" Then GA2.LoadPOP "pop2.txt"
    GA1.SelectionMode = SelWheel
    GA2.SelectionMode = SelWheel
    For I = 1 To NH
        BR1.TransferGAGenesToBrain GA1, I
        BR2.TransferGAGenesToBrain GA2, I
        GA1.IndiFitness(I) = StartFit
        GA2.IndiFitness(I) = StartFit
    Next

    InitShots

    '    TimerMain.Enabled = True

    pPop1.Width = NH * 2
    pPOP2.Width = NH
    pPop1.Height = BR1.GetNofTotalGenes
    pPOP2.Height = BR2.GetNofTotalGenes

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End

End Sub



'Private Sub TimerMain_Timer()
Public Sub MAINLOOP()
    Dim C              As Long

    Dim II             As Long

    Dim X              As Integer
    Dim Y              As Integer

    Dim X0             As Long
    Dim Y0             As Long

    Dim x1             As Long
    Dim y1             As Long
    Dim x2             As Long
    Dim y2             As Long
    Dim Xto            As Single
    Dim Yto            As Single

    Dim UpScia         As Boolean
    Dim TmpDoScia      As Boolean

    Do

        UpScia = IIf(CurTime Mod SciaFreq = 0, True, False)
        TmpDoScia = DoScia And UpScia
        For I = 1 To NH
    '    If I = IndexBest1 Then Stop
        
            H1(I).Move TmpDoScia
            H2(I).Move TmpDoScia
        Next
        NotOverlap

        'PIC.Cls
        BitBlt PIC.hdc, 0, 0, PIC.ScaleWidth, PIC.ScaleHeight, PIC.hdc, 0, 0, vbBlack    'ness


        If chCAM Then
            If scrollH = 0 Then
                If ScrollTEAM = 1 Then
                    Xto = H1(IndexBest1).PosX
                    Yto = H1(IndexBest1).PosY
                Else
                    Xto = H2(IndexBest2).PosX
                    Yto = H2(IndexBest2).PosY
                End If
            Else
                If ScrollTEAM = 1 Then
                    Xto = H1(scrollH).PosX
                    Yto = H1(scrollH).PosY
                Else
                    Xto = H2(scrollH).PosX
                    Yto = H2(scrollH).PosY
                End If
            End If
            PanX = PanX * 0.9875 + Xto * 0.0125
            PanY = PanY * 0.9875 + Yto * 0.0125
            If (CurTime Mod 100) >= 38 Then MyCircle PIC.hdc, XtoScreen(Xto), YtoScreen(Yto), 20 * ZOOM, 1, vbYellow
        End If





        C = RGB(100, 100, 100)
        Y0 = YtoScreen(0)
        y2 = YtoScreen(MaxY)
        For x1 = 0 To MaxX Step 100
            x2 = XtoScreen(x1 \ 1)
            FastLine PIC.hdc, x2, Y0, x2, y2, 1, C
        Next
        X0 = XtoScreen(0)
        x2 = XtoScreen(MaxX)
        For y1 = 0 To MaxY Step 100
            y2 = YtoScreen(y1 \ 1)
            FastLine PIC.hdc, X0, y2, x2, y2, 1, C
        Next


        x1 = XtoScreen(0)
        y1 = YtoScreen(0)
        x2 = XtoScreen(MaxX \ 1)
        y2 = YtoScreen(MaxY \ 1)
        FastLine PIC.hdc, x1, y1, x2, y1, 1, vbWhite
        FastLine PIC.hdc, x1, y2, x2, y2, 1, vbWhite
        FastLine PIC.hdc, x1, y1, x1, y2, 1, vbWhite
        FastLine PIC.hdc, x2, y1, x2, y2, 1, vbWhite


        For I = 1 To NH
     '   If I = IndexBest1 Then Stop
        
            H1(I).DRAW PIC.hdc, StartFit - GA1.IndiFitness(I), 1, DoScia
            H2(I).DRAW PIC.hdc, StartFit - GA2.IndiFitness(I), 2, DoScia
        Next
        H1(IndexBest1).DRAW PIC.hdc, StartFit - GA1.IndiFitness(IndexBest1), 1, DoScia, True
        H2(IndexBest2).DRAW PIC.hdc, StartFit - GA2.IndiFitness(IndexBest2), 2, DoScia, True


        MoveAndDrawShots PIC.hdc, ZOOM


        ComputeEnemiesDistances

        RunBrains

        ' For I = 1 To NH
        '        PIC.ForeColor = vbRed
        '        For II = 1 To GA1.NumberOfIndivid
        '            PIC.CurrentX = H1(II).PosX * ZOOM
        '            PIC.CurrentY = H1(II).PosY * ZOOM
        '            PIC.Print GA1.IndiFitness(II)
        '        Next
        '        PIC.ForeColor = vbCyan
        '        For II = 1 To GA2.NumberOfIndivid
        '            PIC.CurrentX = H2(II).PosX * ZOOM
        '            PIC.CurrentY = H2(II).PosY * ZOOM
        '            PIC.Print GA2.IndiFitness(II)
        '        Next
        '    Next
        PIC.Refresh
        DoEvents




        'ReChargecollision
        'HUMANcollision

                
        CurTime = CurTime + 1

        If CurTime Mod 100 = 0 Then

            AVG1 = GA1.ComputeAVGfit
            AVG2 = GA2.ComputeAVGfit


            ' PicWin.Cls
            BitBlt PicWin.hdc, 0, 0, PicWin.ScaleWidth, PicWin.ScaleHeight, PicWin.hdc, 0, 0, vbBlackness

            If AVG1 > AVG2 Then
                C = 1 * (AVG1 - AVG2)
                C = RGB(C * 0.5, C, C)
            Else
                C = 1 * (AVG2 - AVG1)
                C = RGB(C, C * 0.5, C * 0.5)
            End If


            PicWin.Line (PicWin.Width \ 2, 0)-(PicWin.Width \ 2 + KpicWin * (AVG1 - AVG2), PicWin.Width), C, BF

            lAVG1 = AVG1 & " " & (AVG2 - AVG1)
            lAVG1.BackColor = IIf(AVG1 < AVG2, vbWhite, RGB(200, 200, 200))
            lAVG2 = AVG2 & " " & (AVG1 - AVG2)
            lAVG2.BackColor = IIf(AVG2 < AVG1, vbWhite, RGB(200, 200, 200))


            Y = PY
            X = (AVG1 - AVG2) * 0.15 + PicRes.Width * 0.5
            PicRes.Line (oX, oY)-(X, Y), vbBlue
            oX = X
            oY = Y
            PY = PY + (PicRes.Height / (MatchLenght * 0.01)) '/66=*0.01515

            Me.Caption = CurTime

            If CurTime > MatchLenght Or Abs(AVG1 - AVG2) > MaxFitDist Then
            
                GENES
                CurTime = 0
                InitShots
                For II = 1 To UBound(H1)
                    H1(II).LastShotTime = 0
                    H1(II).HitTIME = 0
                    H1(II).CantShot = False
                Next
                For II = 1 To UBound(H2)
                    H2(II).LastShotTime = 0
                    H2(II).HitTIME = 0
                    H2(II).CantShot = False
                Next
                PY = 0
                oX = PicRes.Width * 0.5
                oY = 0
                PicRes.Cls
                PicRes.Line (PicRes.Width / 2, 0)-(PicRes.Width / 2, PicRes.Height), vbBlack

            End If
        Else

            ' If CurTime Mod 1500 = 0 Then GENES
        End If


        'Stop
        'Me.Refresh
        DoEvents

    Loop While True


End Sub



Public Sub RunBrains()
    Dim I              As Long
    Dim NearToTeam1    As Long
    Dim NearToTeam2    As Long
    Dim A1             As Single
    Dim A2             As Single
    Dim D1             As Single
    Dim D2             As Single



    Dim rT1            As Long
    Dim rT2            As Long
    Dim ShootByEN2     As Long
    Dim ShootByEN1     As Long
    Dim EN2sIDX        As Long
    Dim EN1sIDX        As Long

    Dim SD1            As Single
    Dim SD2            As Single
    Dim SA1            As Single
    Dim SA2            As Single

    Dim AR             As Single
    Dim AL             As Single

        

    Dim Inp1(1 To NofInputs) As Double
    Dim Out1()         As Double
    Dim Inp2(1 To NofInputs) As Double
    Dim Out2()         As Double



    Dim H              As Integer
    Dim w              As Integer
    Dim E              As Long

    Dim KDG            As Single
    

     
    H = PicInput.Height


    For I = 1 To NH

        If chFIXEnemy.Value = vbChecked Then

            If H1(I).EnemyToAttack = 0 Or H2(H1(I).EnemyToAttack).Hitten Then
                NearToTeam1 = NearestEnemy(1, I): H1(I).EnemyToAttack = NearToTeam1
            Else
                NearToTeam1 = H1(I).EnemyToAttack
            End If

            If H2(I).EnemyToAttack = 0 Or H1(H2(I).EnemyToAttack).Hitten Then
                NearToTeam2 = NearestEnemy(2, I): H2(I).EnemyToAttack = NearToTeam2
            Else
                NearToTeam2 = H2(I).EnemyToAttack
            End If
        Else
            NearToTeam1 = NearestEnemy(1, I): H1(I).EnemyToAttack = NearToTeam1
            NearToTeam2 = NearestEnemy(2, I): H2(I).EnemyToAttack = NearToTeam2
        End If

        '        Stop


        A1 = ComputeAngle(1, I, NearToTeam1)
        A2 = ComputeAngle(2, I, NearToTeam2)

        D1 = Sqr(EnemyDist(I, NearToTeam1))
        D2 = Sqr(EnemyDist(I, NearToTeam2))


        AngToLeftRight A1, AL, AR
        Inp1(1) = AL
        Inp1(2) = AR
        Inp1(3) = 1 - D1 / MaxEYEdist
        If Inp1(3) < 0 Then Inp1(3) = 0

        AngToLeftRight A2, AL, AR
        Inp2(1) = AL
        Inp2(2) = AR
        Inp2(3) = 1 - D2 / MaxEYEdist
        If Inp2(3) < 0 Then Inp2(3) = 0


        '   Stop

        '-------------------------
        SD1 = NearestShot(1, I, ShootByEN2, EN2sIDX)
        SD2 = NearestShot(2, I, ShootByEN1, EN1sIDX)


        If I = IndexBest1 And Sho(2, ShootByEN2, EN2sIDX).Enabled = True Then
            'MyCircle PIC.Hdc, XtoScreen(H1(I).PosX), YtoScreen(H1(I).PosY), 10 * ZOOM, 1, vbRed
            MyCircle PIC.hdc, XtoScreen(Sho(2, ShootByEN2, EN2sIDX).X), YtoScreen(Sho(2, ShootByEN2, EN2sIDX).Y), 8 * ZOOM, 1, vbYellow
        End If


        SA1 = Atan2(H1(I).PosX - Sho(2, ShootByEN2, EN2sIDX).X, H1(I).PosY - Sho(2, ShootByEN2, EN2sIDX).Y)
        SA1 = AngleDiff01(H1(I).ANG, SA1)

        SA2 = Atan2(H2(I).PosX - Sho(1, ShootByEN1, EN1sIDX).X, H2(I).PosY - Sho(1, ShootByEN1, EN1sIDX).Y)
        SA2 = AngleDiff01(H2(I).ANG, SA2)

        AngToLeftRight SA1, AL, AR
        Inp1(4) = AL
        Inp1(5) = AR
        Inp1(6) = 1 - SD1 / MaxEYEdist
        If Inp1(6) < 0 Then Inp1(6) = 0
        AngToLeftRight SA2, AL, AR
        Inp2(4) = AL
        Inp2(5) = AR
        Inp2(6) = 1 - SD2 / MaxEYEdist
        If Inp2(6) < 0 Then Inp2(6) = 0


        If SD1 < ((ShotSpeed + MaxHSpeed) * 2 + 8) Then
            If CurTime - H1(I).HitTIME > FrozenTime Then
                GA1.IndiFitness(I) = GA1.IndiFitness(I) - HittenPTS    '80
                H1(I).HitTIME = CurTime
                'Stop

                H1(I).VxHit = H1(I).VxHit + Sho(2, ShootByEN2, EN2sIDX).vX
                H1(I).VyHit = H1(I).VyHit + Sho(2, ShootByEN2, EN2sIDX).vY
                H1(I).ANGVel = H1(I).ANGVel + (Rnd - 0.5) * 3
                
                                'H2(ShootByEN2).ANGVel = H2(ShootByEN2).ANGVel + Rnd - 0.5

                'If ShootByEN2 = NearToTeam1 Then
                If Sho(2, ShootByEN2, EN2sIDX).TE = NearToTeam2 Then

                    GA2.IndiFitness(ShootByEN2) = GA2.IndiFitness(ShootByEN2) - HitPTS * 2
                    'H1(I).PosX = MaxX * 0.25
                    'H1(I).PosY = Rnd * MaxY
                    'H1(I).ANG = Rnd * 0.2 - 0.1    'Rnd * pi2
                    'H1(I).Vel = 0    '4
                Else
                    GA2.IndiFitness(ShootByEN2) = GA2.IndiFitness(ShootByEN2) - HitPTS
                End If

                Sho(2, ShootByEN2, EN2sIDX).Enabled = False
                H2(ShootByEN2).AvailShots = H2(ShootByEN2).AvailShots + 1
            End If
        End If
        If SD2 < ((ShotSpeed + MaxHSpeed) * 2 + 8) Then
            If CurTime - H2(I).HitTIME > FrozenTime Then
                GA2.IndiFitness(I) = GA2.IndiFitness(I) - HittenPTS    '80
                H2(I).HitTIME = CurTime

                H2(I).VxHit = H2(I).VxHit + Sho(1, ShootByEN1, EN1sIDX).vX
                H2(I).VyHit = H2(I).VyHit + Sho(1, ShootByEN1, EN1sIDX).vY
                H2(I).ANGVel = H2(I).ANGVel + (Rnd - 0.5) * 3
                
                'H1(ShootByEN1).ANGVel = H1(ShootByEN1).ANGVel + Rnd - 0.5
                'If ShootByEN1 = NearToTeam2 Then
                If Sho(1, ShootByEN1, EN1sIDX).TE = NearToTeam1 Then
                    GA1.IndiFitness(ShootByEN1) = GA1.IndiFitness(ShootByEN1) - HitPTS * 2    '101
                    'H2(I).PosX = MaxX * 0.75
                    'H2(I).PosY = Rnd * MaxY
                    'H2(I).ANG = -PI + Rnd * 0.2 - 0.1    'Rnd * pi2
                    'H2(I).Vel = 0    '4
                Else
                    GA1.IndiFitness(ShootByEN1) = GA1.IndiFitness(ShootByEN1) - HitPTS
                End If

                Sho(1, ShootByEN1, EN1sIDX).Enabled = False
                H1(ShootByEN1).AvailShots = H1(ShootByEN1).AvailShots + 1
            End If
        End If

        '-------------------------------

        Inp1(7) = H1(I).AvailShots / MaxShots
        Inp2(7) = H2(I).AvailShots / MaxShots

        Inp1(8) = AngleDiff01(H1(I).ANG, H2(NearToTeam1).ANG)
        Inp2(8) = AngleDiff01(H2(I).ANG, H1(NearToTeam2).ANG)

        Inp1(9) = H1(I).SPEED / MaxHSpeed
        Inp2(9) = H2(I).SPEED / MaxHSpeed

        Inp1(10) = H2(NearToTeam1).SPEED / MaxHSpeed
        Inp2(10) = H1(NearToTeam2).SPEED / MaxHSpeed

        'Inp1(11) = H2(NearToTeam1).ACC + 0.25
        'Inp2(11) = H1(NearToTeam2).ACC + 0.25

        A1 = H1(I).DirMoving
        A2 = H2(I).DirMoving

        Inp1(11) = 1 - AngleDiff01(A1, H2(NearToTeam1).DirMoving)
        Inp2(11) = 1 - AngleDiff01(A2, H1(NearToTeam2).DirMoving)


        'hilights nearset enemy of best of pop1
        If I = IndexBest1 Then MyCircle PIC.hdc, XtoScreen(H2(NearToTeam1).PosX), YtoScreen(H2(NearToTeam1).PosY), 11 * ZOOM, 1, vbYellow
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Stop
If CurTime Mod 30 = 0 Then
        Out1 = BR1.RUN(I, 1, Inp1)
        Out2 = BR2.RUN(I, 1, Inp2)
        
        'Choose wich brain cell to run to determinate the movement
        If Out1(1) > Out1(2) Then
            H1(I).BRtoRUN = 2
            Else
            H1(I).BRtoRUN = 3
        End If
       
        If Out2(1) > Out2(2) Then
            H2(I).BRtoRUN = 2
            Else
            H2(I).BRtoRUN = 3
        End If
End If

        Out1 = BR1.RUN(I, H1(I).BRtoRUN, Inp1)
        Out2 = BR2.RUN(I, H2(I).BRtoRUN, Inp2)
        
         If I = IndexBest1 Then
            w = PicInput.Width \ NofInputs
            For E = 1 To NofInputs
                PicInput.Line ((E - 1) * w, 0)-(E * w, H), RGB(Inp1(E) * 255, 0, 0), BF
            Next

            w = PicInput.Width \ 3
            For E = 1 To 3
                PicOUT.Line ((E - 1) * w, 0)-(E * w, H), RGB(0, Out1(E) * 255, 0), BF
            Next

            If CurTime Mod 10 = 0 Then BR1.DRAW IndexBest1, H1(IndexBest1).BRtoRUN: PicN.Refresh
        End If



'        H1(I).ANGVel = H1(I).ANGVel - Out1(1) * 0.01
'        H1(I).ANGVel = H1(I).ANGVel + Out1(2) * 0.01
'        H1(I).ACC = Out1(3) - 0.25    ' H1(I).ACC + 0.1 * (Out1(3) - 0.25)
'        H2(I).ANGVel = H2(I).ANGVel - Out2(1) * 0.01
'        H2(I).ANGVel = H2(I).ANGVel + Out2(2) * 0.01
'        H2(I).ACC = Out2(3) - 0.25    ' H2(I).ACC + 0.1 * (Out2(3) - 0.25)

        
        H1(I).ACCLeft = Out1(1) - 0.25
        H1(I).ACCRight = Out1(2) - 0.25
        H2(I).ACCLeft = Out2(1) - 0.25
        H2(I).ACCRight = Out2(2) - 0.25
        


        If (Out1(3) > 0.5) And (H1(I).AvailShots > 0) And (CurTime - H1(I).LastShotTime > ShotDelay) Then
            If H1(I).CantShot = False Then FireShot 1, I, NearToTeam1: GA1.IndiFitness(I) = GA1.IndiFitness(I) - FirePTS
        End If

        If (Out2(3) > 0.5) And (H2(I).AvailShots > 0) And (CurTime - H2(I).LastShotTime > ShotDelay) Then
            If H2(I).CantShot = False Then FireShot 2, I, NearToTeam2: GA2.IndiFitness(I) = GA2.IndiFitness(I) - FirePTS
        End If


    Next


End Sub


'Public Sub ReChargecollision()
'Dim I         As Long
'Dim J         As Long
'
'    For I = 1 To NH
'        For J = 1 To NF
'            'If Distance(H(I).PosX, H(I).PosY, F(J).PosX, F(J).PosY) < (H(I).R + F(J).R) Then
'
'            If Abs(H(I).PosX - F(J).PosX) < F(J).R Then
'                If Abs(H(I).PosY - F(J).PosY) < F(J).R Then
'                    F(J).PosX = Rnd * MaxX
'                    F(J).PosY = Rnd * MaxY'
'
'                    If H(I).Vel > 0.05 Then GA.IndiFitness(I) = GA.IndiFitness(I) - 10    'H(I).Vel * 5
'                End If
'
'            End If
'        Next
'    Next
'
'End Sub
'------------------------------------------------------------------------------------
'Public Sub HUMANcollision()
'Dim I         As Long
'Dim J         As Long'

'    For I = 1 To NH - 1
'       For J = I + 1 To NH
'If Distance(H(I).PosX, H(I).PosY, H(J).PosX, H(J).PosY) < 25 Then
'
'           If Abs(H(I).PosX - H(J).PosX) < 10 Then
'              If Abs(H(I).PosY - H(J).PosY) < 10 Then
'
'
'                    If H(I).Vel = H(J).Vel Then
'                       GA.IndiFitness(I) = GA.IndiFitness(I) + H(I).Vel * 1
'                    GA.IndiFitness(J) = GA.IndiFitness(J) + H(J).Vel * 1
'                      H(I).PosX = Rnd * MaxX
'                     H(I).PosY = Rnd * MaxY
'                    H(J).PosX = Rnd * MaxX
'                   H(J).PosY = Rnd * MaxY
'              Else
'                 If H(I).Vel > H(J).Vel Then
'                           GA.IndiFitness(I) = GA.IndiFitness(I) + H(I).Vel * 1
'                           GA.IndiFitness(J) = GA.IndiFitness(J) + H(J).Vel * 0.1
'
'                            H(I).PosX = Rnd * MaxX
'                            H(I).PosY = Rnd * MaxY
'                        Else
'                            GA.IndiFitness(J) = GA.IndiFitness(J) + H(J).Vel * 1
'                            GA.IndiFitness(I) = GA.IndiFitness(I) + H(I).Vel * 0.1
'
''
'                           H(J).PosX = Rnd * MaxX
'                           H(J).PosY = Rnd * MaxY
''                      End If
'                  End If
''
'
'               End If
'
'            End If
'        Next
'    Next
'
'
'End Sub

Public Sub GENES()
    Dim X              As Long
    Dim Y              As Long
    Dim KDG            As Single

    For I = 1 To NH

        H1(I).Walked = 0
        H2(I).Walked = 0

    Next

    If GA1.ComputeAVGfit <= GA2.ComputeAVGfit Then
        Win1 = Win1 + 1
    Else
        Win2 = Win2 + 1
    End If

    Lwin1 = "RED :" & Win1
    Lwin2 = "CYAN:" & Win2

    'EVOLVE LOSER
    If GA1.ComputeAVGfit > GA2.ComputeAVGfit Then
        Lprewin = "Previous winner: CYAN"
        GA1.EVOLVE2
        
        For I = 1 To NH
            GA1.IndiFitness(I) = StartFit
        Next
        For I = 1 To NH
            GA2.IndiFitness(I) = GA2.IndiFitness(I) - (GA2.STAT_AVGfit - StartFit)
        Next

        If Rnd < 0.1 Then
            GA2.EVOLVE2
            For I = 1 To NH
                GA2.IndiFitness(I) = StartFit
            Next
        End If

    Else
        GA2.EVOLVE2
        For I = 1 To NH
            GA2.IndiFitness(I) = StartFit
        Next
        For I = 1 To NH
            GA1.IndiFitness(I) = GA1.IndiFitness(I) - (GA1.STAT_AVGfit - StartFit)
        Next

        If Rnd < 0.1 Then
            GA1.EVOLVE2
            For I = 1 To NH
                GA1.IndiFitness(I) = StartFit
            Next
        End If

        Lprewin = "Previous winner: RED"
    End If
    'GA1.COMPUTEGENES
    'GA2.COMPUTEGENES


    GEN.Text = GA1.sTAT_NofGEN
    ACC.Text = GA1.STAT_NofACC
    MUT.Text = GA1.STAT_NofMUT
    NEWr.Text = GA1.STAT_NofNEW
    gAVG.Text = Int(GA1.STAT_AVGfit * 100) / 100
    BFIT = GA1.STAT_GenerBestFit

    Text1.Text = GA2.sTAT_NofGEN
    Text2.Text = GA2.STAT_NofACC
    Text3.Text = GA2.STAT_NofMUT
    Text4.Text = GA2.STAT_NofNEW
    Text6.Text = Int(GA2.STAT_AVGfit * 100) / 100
    Text5.Text = GA2.STAT_GenerBestFit
    DoEvents

    IndexBest1 = GA1.STAT_GenerBestFitINDX
    IndexBest2 = GA2.STAT_GenerBestFitINDX

    For I = 1 To NH

        BR1.TransferGAGenesToBrain GA1, I
        BR2.TransferGAGenesToBrain GA2, I

        With H1(I)
            .PosX = MaxX * 0.25
            .PosY = Rnd * MaxY
            .vX = 0
            .vY = 0
            .ANG = Rnd * 0.2 - 0.1    'Rnd * pi2
            .ACCLeft = 0.5       '4
            .ACCRight = 0.5       '4
            .AvailShots = MaxShots
            .Hitten = False
            .HitTIME = 0
        End With
        With H2(I)
            .PosX = MaxX * 0.75
            .PosY = Rnd * MaxY
            .vX = 0
            .vY = 0
            .ANG = -PI + Rnd * 0.2 - 0.1    'Rnd * pi2
            .ACCLeft = 0.5       '4
            .ACCRight = 0.5       '4
            .AvailShots = MaxShots
            .Hitten = False
            .HitTIME = 0
        End With

    Next


    GA1.SavePOP "Pop1.txt"
    GA2.SavePOP "Pop2.txt"

    ' For I = 1 To NH
    '     GA1.IndiFitness(I) = StartFit
    '     GA2.IndiFitness(I) = StartFit
    ' Next


    InitShots

    DoEvents

    KDG = 255 / (GA1.GeneValuesMax - GA1.GeneValuesMin)


    For X = 0 To NH - 1
        For Y = 0 To GA1.NofGenesMAX - 1
            SetPixel pPop1.hdc, X, Y, RGB(GA1.GeneValue(X + 1, Y + 1) * KDG, 0, 0)
            SetPixel pPop1.hdc, X + NH, Y, RGB(0, GA2.GeneValue(X + 1, Y + 1) * KDG, GA2.GeneValue(X + 1, Y + 1) * KDG)

            SetPixel pPOP2.hdc, X, Y, RGB(0, GA2.GeneValue(X + 1, Y + 1) * KDG, GA2.GeneValue(X + 1, Y + 1) * KDG)

        Next
    Next
    pPop1.Refresh
    pPOP2.Refresh

    DoEvents

 '   SavePicture pPop1.Image, App.Path & "\Frames\" & Format(GA1.sTAT_NofGEN + GA2.sTAT_NofGEN, "00000000") & ".bmp"




End Sub

