Attribute VB_Name = "modRENDER"
Public PanX            As Single
Public PanY            As Single
Public CenX            As Single
Public CenY            As Single

Public PanZoomChanged  As Boolean
Public Navigating      As Boolean

Public MaxXPic         As Long
Public MaxYpic         As Long

Public MaxX            As Single
Public MaxY            As Single


Public ZOOM            As Single
Public InvZoom As Single


Public CAMFollowing    As Boolean

Public DoScia As Boolean

Public Function XtoScreen(X As Single) As Long
    XtoScreen = ZOOM * (X - PanX) + CenX
End Function
Public Function YtoScreen(Y As Single) As Long
    YtoScreen = ZOOM * (Y - PanY) + CenY
End Function

Public Function xfromScreen(X As Long) As Single
    xfromScreen = (X - CenX) * InvZoom + PanX
End Function
Public Function yfromScreen(Y As Long) As Single
    yfromScreen = (Y - CenY) * InvZoom + PanY
End Function
Public Function IsInsideScreen(X As Long, Y As Long) As Boolean
' IsInsideScreen = False

    If X < 0 Then Exit Function
    If X > MaxXPic Then Exit Function
    If Y < 0 Then Exit Function
    If Y > MaxYpic Then Exit Function

    IsInsideScreen = True

End Function
