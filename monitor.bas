Attribute VB_Name = "Module9"
Option Explicit
Private Declare Function EnumDisplayMonitors Lib "User32" (ByVal hdc As Long, lprcClip As Any, ByVal lpfnEnum As Long, dwData As Any) As Long
Public Declare Function MonitorFromPoint Lib "User32" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
Private Declare Function MonitorFromWindow Lib "User32" (ByVal hWND As Long, ByVal dwFlags As Long) As Long

Private Declare Function GetMonitorInfo Lib "User32" Alias "GetMonitorInfoA" (ByVal hmonitor As Long, ByRef lpmi As MONITORINFO) As Long
Private Declare Function GetWindowRect Lib "User32" (ByVal hWND As Long, lpRect As RECT) As Long
Private Declare Function UnionRect Lib "User32" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long
Private Declare Function OffsetRect Lib "User32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveWindow Lib "User32" (ByVal hWND As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'Private Type RECT
 '   Left As Long
  '  Top As Long
   ' Right As Long
    'Bottom As Long
'End Type
Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type
Public Const MONITORINFOF_PRIMARY = &H1
Public Const MONITOR_DEFAULTTONEAREST = &H2
Public Const MONITOR_DEFAULTTONULL = &H0
Public Const MONITOR_DEFAULTTOPRIMARY = &H1
Dim rcMonitors() As RECT 'coordinate array for all monitors
Dim rcVS         As RECT 'coordinates for Virtual Screen

Public Type Screens
    Top As Long
    Left As Long
    Height As Long
    width As Long
    primary As Boolean
    handler As Long
End Type
Public ScrInfo() As Screens
Public Console As Long
' MI.dwFlags = MONITORINFOF_PRIMARY

Private Const SM_CXVIRTUALSCREEN = 78
Private Const SM_CYVIRTUALSCREEN = 79
Private Const SM_CMONITORS = 80
Private Const SM_SAMEDISPLAYFORMAT = 81

Private Declare Function GetSystemMetrics Lib "User32" ( _
   ByVal nIndex As Long) As Long

Public Property Get VirtualScreenWidth() As Long
   VirtualScreenWidth = (GetSystemMetrics(SM_CXVIRTUALSCREEN) + 2000) * dv15
End Property
Public Property Get VirtualScreenHeight() As Long
   VirtualScreenHeight = (GetSystemMetrics(SM_CYVIRTUALSCREEN) + 100) * dv15
End Property
Public Property Get DisplayMonitorCount() As Long
   DisplayMonitorCount = GetSystemMetrics(SM_CMONITORS)
End Property
Public Property Get AllMonitorsSame() As Long
   AllMonitorsSame = GetSystemMetrics(SM_SAMEDISPLAYFORMAT)
End Property
Public Sub GetMonitorsNow()
  Dim N As Long
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, N
    Console = FindMonitorFromMouse
End Sub
Function EnumMonitors(F As Form) As Long
    Dim N As Long
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, N
    With F
        .Move .Left, .Top, (rcVS.Right - rcVS.Left) * 2 + .width - .ScaleWidth, (rcVS.Bottom - rcVS.Top) * 2 + .Height - .ScaleHeight
    End With
    F.Scale (rcVS.Left, rcVS.Top)-(rcVS.Right, rcVS.Bottom)
    F.Caption = N & " Monitor" & IIf(N > 1, "s", vbNullString)
    F.lblMonitors(0).Appearance = 0 'Flat
    F.lblMonitors(0).BorderStyle = 1 'FixedSingle
    For N = 0 To N - 1
        If N Then
            Load F.lblMonitors(N)
            F.lblMonitors(N).Visible = True
        End If
        With rcMonitors(N)
            F.lblMonitors(N).Move .Left, .Top, .Right - .Left, .Bottom - .Top
            F.lblMonitors(N).Caption = "Monitor " & N + 1 & vbLf & _
                .Right - .Left & " x " & .Bottom - .Top & vbLf & _
                "(" & .Left & ", " & .Top & ")-(" & .Right & ", " & .Bottom & ")"
        End With
    Next
End Function
Private Function MonitorEnumProc(ByVal hmonitor As Long, ByVal hdcMonitor As Long, lprcMonitor As RECT, dwData As Long) As Long
    Dim mi As MONITORINFO
    ReDim Preserve rcMonitors(dwData)
    ReDim Preserve ScrInfo(dwData)
    rcMonitors(dwData) = lprcMonitor
    mi.cbSize = Len(mi)
    GetMonitorInfo hmonitor, mi
    
    With ScrInfo(dwData)
    .Left = mi.rcMonitor.Left * dv15
    .Top = mi.rcMonitor.Top * dv15
    If IsWine Then
    .Height = (mi.rcMonitor.Bottom - mi.rcMonitor.Top + 1) * dv15
    .width = (mi.rcMonitor.Right - mi.rcMonitor.Left + 1) * dv15
    Else
    
    .Height = (mi.rcMonitor.Bottom - mi.rcMonitor.Top + 1) * dv15
    .width = (mi.rcMonitor.Right - mi.rcMonitor.Left + 1) * dv15
    End If
    .primary = CBool(mi.dwFlags = MONITORINFOF_PRIMARY)
    .handler = hmonitor
    End With
    UnionRect rcVS, rcVS, lprcMonitor 'merge all monitors together to get the virtual screen coordinates
    dwData = dwData + 1 'increase monitor count
    MonitorEnumProc = 1 'continue
End Function

Sub SavePosition(hWND As Long)
    Dim rc As RECT
    GetWindowRect hWND, rc 'save position in pixel units
    SaveSetting "Multi Monitor Demo", "Position", "Left", rc.Left
    SaveSetting "Multi Monitor Demo", "Position", "Top", rc.Top
End Sub


Function FindPrimary() As Long
Dim i As Long
For i = 0 To UBound(ScrInfo())
If ScrInfo(i).primary Then FindPrimary = i: Exit Function
Next i
End Function
Function FindFormSScreenCorner(Z As Object)
Dim F As Form
If TypeOf Z Is Form Then
Set F = Z
Else
Set F = Z.Parent
End If
On Error Resume Next
Dim thismonitor As Long
thismonitor = FindMonitorFromPixel(F.Left, F.Top)
Dim i As Long
For i = 0 To UBound(ScrInfo())
If thismonitor = ScrInfo(i).handler Then FindFormSScreenCorner = i:   Exit Function
Next i
End Function
Function FindFormSScreen(Z As Object)
Dim F As Form
If TypeOf Z Is Form Then
Set F = Z
Else
Set F = Z.Parent
End If
On Error Resume Next
Dim thismonitor As Long
thismonitor = MonitorFromWindow(F.hWND, MONITOR_DEFAULTTONEAREST)
Dim i As Long
For i = 0 To UBound(ScrInfo())
If thismonitor = ScrInfo(i).handler Then FindFormSScreen = i:   Exit Function
Next i
End Function
Function FindMonitorFromPixel(x, y) As Long
Dim x1 As Long, y1 As Long
x1 = x \ dv15
y1 = y \ dv15
Dim i As Long
For i = 0 To UBound(ScrInfo())
If ScrInfo(i).handler = MonitorFromPoint(x1, y1, MONITOR_DEFAULTTONEAREST) Then FindMonitorFromPixel = i: Exit Function
Next i

End Function
Function FindMonitorFromMouse()
'
   ' - offset
Dim x As Long, y As Long, tp As POINTAPI
GetCursorPos tp
x = tp.x
y = tp.y
Dim i As Long
For i = 0 To UBound(ScrInfo())
If ScrInfo(i).handler = MonitorFromPoint(x, y, MONITOR_DEFAULTTONEAREST) Then FindMonitorFromMouse = i: Exit Function
Next i
End Function
Sub MoveFormToOtherMonitor(F As Form)
Dim k As Long, Z As Long
'k = FindMonitorFromPixel(F.Left, F.Top)
Z = FindMonitorFromMouse
'If k <> Z Then
' center to z
If F.width > ScrInfo(Z).width Then
    If F.Height > ScrInfo(Z).Height Then
        F.Move ScrInfo(Z).Left, ScrInfo(Z).Top
    Else
        F.Move ScrInfo(Z).Left, ScrInfo(Z).Top + (ScrInfo(Z).Height - F.Height) / 2
    End If
    
ElseIf F.Height > ScrInfo(Z).Height Then
    F.Move ScrInfo(Z).Left + (ScrInfo(Z).width - F.width) / 2, ScrInfo(Z).Top
Else
 ' F.Move ScrInfo(Z).Left + (ScrInfo(Z).width - F.width) / 2, ScrInfo(Z).Top + (ScrInfo(Z).Height - F.Height) / 2

End If
'End If
End Sub
Sub MoveFormToOtherMonitorOnly(F As Form, Optional flag As Boolean)
Dim k As Long, Z As Long
Dim nowX As Long, nowY As Long
k = FindMonitorFromPixel(F.Left, F.Top)
Z = FindMonitorFromMouse
If k = Z Then
If flag Then
Dim tp As POINTAPI
GetCursorPos tp
nowX = tp.x * dv15
nowY = tp.y * dv15
flag = False
Else
flag = False
nowX = F.Left - ScrInfo(k).Left + ScrInfo(Z).Left
nowY = F.Top - ScrInfo(k).Top + ScrInfo(Z).Top
'Exit Sub
End If
Else
nowX = F.Left - ScrInfo(k).Left + ScrInfo(Z).Left
nowY = F.Top - ScrInfo(k).Top + ScrInfo(Z).Top
End If

If nowX > ScrInfo(Z).Left + ScrInfo(Z).width Then
    nowX = ScrInfo(Z).Left + ScrInfo(Z).width * 2 / 3
End If
If nowX + F.width > ScrInfo(Z).Left + ScrInfo(Z).width Then
    If F.width < ScrInfo(Z).width Then
    nowX = ScrInfo(Z).Left + ScrInfo(Z).width - F.width
    Else
    nowX = ScrInfo(Z).Left
    End If
End If
If nowY > ScrInfo(Z).Top + ScrInfo(Z).Height Then
    nowY = ScrInfo(Z).Top + ScrInfo(Z).Height * 2 / 3
End If
If nowY + F.Height > ScrInfo(Z).Top + ScrInfo(Z).Height Then
    If F.Height < ScrInfo(Z).Height Then
    nowY = ScrInfo(Z).Top + ScrInfo(Z).Height - F.Height
    Else
    nowY = ScrInfo(Z).Top
    End If
End If

If F.width > ScrInfo(Z).width Then
    If F.Height > ScrInfo(Z).Height Then
        nowX = ScrInfo(Z).Left
        nowY = ScrInfo(Z).Top
    Else
        nowX = ScrInfo(Z).Left
        nowY = ScrInfo(Z).Top + (ScrInfo(Z).Height - F.Height) / 2
    End If
    
ElseIf F.Height > ScrInfo(Z).Height Then
    nowX = ScrInfo(Z).Left + (ScrInfo(Z).width - F.width) / 2
    nowY = ScrInfo(Z).Top
ElseIf flag Then
    nowX = ScrInfo(Z).Left + (ScrInfo(Z).width - F.width) / 2
    nowY = ScrInfo(Z).Top + (ScrInfo(Z).Height - F.Height) / 2
End If
F.Move nowX, nowY
End Sub
Sub MoveFormToOtherMonitorCenter(F As Form)
Dim k As Long, Z As Long
'k = FindMonitorFromPixel(F.Left, F.Top)
Z = FindMonitorFromMouse
'If k <> Z Then
' center to z
If F.width > ScrInfo(Z).width Then
    If F.Height > ScrInfo(Z).Height Then
        F.Move ScrInfo(Z).Left, ScrInfo(Z).Top
    Else
        F.Move ScrInfo(Z).Left, ScrInfo(Z).Top + (ScrInfo(Z).Height - F.Height) / 2
    End If
    
ElseIf F.Height > ScrInfo(Z).Height Then
    F.Move ScrInfo(Z).Left + (ScrInfo(Z).width - F.width) / 2, ScrInfo(Z).Top
Else
 F.Move ScrInfo(Z).Left + (ScrInfo(Z).width - F.width) / 2, ScrInfo(Z).Top + (ScrInfo(Z).Height - F.Height) / 2

End If
'End If
End Sub
