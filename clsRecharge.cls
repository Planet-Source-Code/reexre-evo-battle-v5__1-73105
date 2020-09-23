VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PosX            As Single
Public PosY            As Single
Public r               As Single

Public Sub DRAW(P As PictureBox, ZOOM As Single)


'P.Circle (PosX * Zoom, PosY * Zoom), R * Zoom, vbGreen

    MyCircle P.Hdc, PosX * ZOOM, PosY * ZOOM, r * ZOOM, 1, vbGreen


End Sub

