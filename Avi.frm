VERSION 5.00
Begin VB.Form AVI 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1065
   ClientLeft      =   6285
   ClientTop       =   -8790
   ClientWidth     =   1380
   Icon            =   "Avi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1065
   ScaleWidth      =   1380
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   285
      Top             =   375
   End
End
Attribute VB_Name = "AVI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim doubleclick As Long
Dim ERNUM As Long



Private Sub Form_Activate()

CLICK_COUNT = CLICK_COUNT + 1
If getout Then Exit Sub
If CLICK_COUNT < 2 Then


If Height > 30 Then
If Not UseAviXY Then
SetWindowPos Me.hWND, HWND_TOPMOST, Me.Left / dv15, _
                        Me.Top / dv15, Me.width / dv15, _
                        Me.Height / dv15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
Else


SetWindowPos Me.hWND, HWND_TOPMOST, aviX / dv15, _
                        aviY / dv15, Me.width / dv15, _
                        Me.Height / dv15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End If
End If
ElseIf AVIRUN Then
getout = True
Else
End If
If getout Then
GETLOST
End If
End Sub

Private Sub Form_Click()
If MediaPlayer1.isMoviePlaying Then GETLOST
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
KeyCode = 0
If Form1.Visible Then
If Form1.TEXT1.Visible Then
Form1.TEXT1.SetFocus
Else
Form1.SetFocus
End If
End If

End Sub

Private Sub Form_Load()
Me.BackColor = Form1.DIS.BackColor
If getout Then Exit Sub
CLICK_COUNT = 0
Dim cd As String
getout = False
On Error Resume Next
If avifile = vbNullString Then
GETLOST
getout = True
Else
MediaPlayer1.hideMovie
MediaPlayer1.FileName = avifile
Timer1.enabled = False

Timer1.Interval = MediaPlayer1.Length

If MediaPlayer1.error = 0 Then

If UseAviSize And MediaPlayer1.Height > 2 Then
If AviSizeX = 0 And AviSizeY = 0 Then
If aviX = 0 And aviY = 0 Then

AviSizeX = (ScrInfo(Console).width - 1) * 0.99
AviSizeY = (ScrInfo(Console).Height - 1) * 0.99
Else
AviSizeX = Me.ScaleWidth
AviSizeY = Me.ScaleHeight
End If
Else
If AviSizeX = 0 And AviSizeY <> 0 Then
AviSizeX = CLng(AviSizeY * MediaPlayer1.width / CDbl(MediaPlayer1.Height))
End If
If AviSizeY = 0 Then
AviSizeY = CLng(AviSizeX * MediaPlayer1.Height / CDbl(MediaPlayer1.width))
End If
If aviX = 0 And aviY = 0 Then

aviX = ((ScrInfo(Console).width - 1) * 0.99 - AviSizeX) / 2 + ScrInfo(Console).Left
aviY = ((ScrInfo(Console).Height - 1) * 0.99 - AviSizeY) / 2 + ScrInfo(Console).Top
End If
End If

Me.Move aviX, aviY, AviSizeX, AviSizeY
MyDoEvents
MediaPlayer1.openMovieWindow Me.hWND, "child"

MediaPlayer1.sizeLocateMovie 0, 0, ScaleX(AviSizeX, vbTwips, vbPixels), ScaleY(AviSizeY, vbTwips, vbPixels) + 1
'Show
ElseIf MediaPlayer1.Height > 2 Then
Me.Move Left, Top, ScaleX(MediaPlayer1.width, vbPixels, vbTwips), ScaleY(MediaPlayer1.Height, vbPixels, vbTwips) + 1

MediaPlayer1.openMovieWindow AVI.hWND, "child"




Else
Me.Move Left, Top, ScaleX(c&, vbPixels, vbTwips), ScaleY(MediaPlayer1.Height, vbPixels, vbTwips)

MediaPlayer1.minimizeMovie
MediaPlayer1.openMovie

End If


If MediaPlayer1.Height <= 2 Then
width = 0
Height = 0
Else

If Not UseAviXY Then
Me.Move ((ScrInfo(Console).width - 1) - width) / 2, ((ScrInfo(Console).Height - 1) - Height) / 2
End If
End If


Timer1.enabled = False



Else
getout = True
width = 0
Height = 0
'Show
End If
End If
AVIUP = True
End Sub

Public Sub Avi2Up()
Timer1.enabled = True
Me.ZOrder
MediaPlayer1.playMovie

Timer1.enabled = True
AVIRUN = True
End Sub



Private Sub Form_Unload(Cancel As Integer)
getout = False
AVIRUN = False
AVIUP = False
End Sub



Public Sub GETLOST()
getout = True
Timer1.enabled = False
Hide
MediaPlayer1.hideMovie
MediaPlayer1.stopMovie
MediaPlayer1.closeMovie
AVIRUN = False
MyDoEvents
If Form1.Visible Then Form1.Refresh: If Form1.DIS.Visible Then Form1.DIS.Refresh
Unload Me
End Sub


Private Sub Frame1_Click()
GETLOST
End Sub

Private Sub Timer1_Timer()

GETLOST

End Sub
