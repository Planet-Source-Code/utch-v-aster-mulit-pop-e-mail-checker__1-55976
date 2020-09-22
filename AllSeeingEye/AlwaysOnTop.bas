Attribute VB_Name = "AlwaysOnTop"
Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2



Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    If SetOnTop Then
    lFlag = HWND_TOPMOST
    Else
    lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub



