VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mutex"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Mymutexs As New FastCollection
Private Const ERROR_ALREADY_EXISTS = 183&
Private Const MUTEX_ALL_ACCESS = &H1F0001
Public LastError As Long
Private Declare Function CreateMutex Lib "KERNEL32" Alias "CreateMutexW" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long

Sub Create(name$)
If Mymutexs.ExistKey(name$) Then exist: Exit Sub
Dim myMutex As Variant
myMutex = CVar(CreateMutex(0, 0, StrPtr(name$)))
LastError = Err.LastDllError
If LastError = 0 Then Mymutexs.AddKey2 name$, myMutex
If LastError = ERROR_ALREADY_EXISTS Then CloseHandle myMutex:    MyEr "Mutex object exist", "Υπάρχει ήδη το αντικείμενο"
End Sub

Private Sub notexist()
MyEr "named mutex not exist", "το αντικείμενο δεν υπάρχει"
End Sub
Private Sub exist()
MyEr "named mutex exist", "το αντικείμενο υπάρχει"
End Sub
Sub Destroy(name$)
If Not Mymutexs.ExistKey(name$) Then notexist: Exit Sub
CloseHandle CLng(Mymutexs.Value)
End Sub

Private Sub Class_Terminate()
Dim i As Long, myMutex As Long
If Mymutexs.count > 0 Then
For i = 0 To Mymutexs.count - 1
Mymutexs.index = i
CloseHandle CLng(Mymutexs.Value)
Next i
End If
End Sub
