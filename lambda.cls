VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "lambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lambda functions in M2000
' in running state copy nonlocal variables to current list
Public name$
Public code$
Private Type Vars
    Value As Variant
    there As Long
End Type
Private Total As Long, fixed As Long, used As Long
Private myVar() As Vars, myHash As New Hash

Sub CopyToVar(bstack As basetask, glb As Boolean, v() As Variant)
Dim i As Long, nm$, where As Long, dummy As Variant, x1 As Long
If used > 0 Then
For i = 0 To used - 1
    myHash.ReadVar i, nm$, where
    If IsObject(myVar(where).Value) Then
    dummy = 0
    myVar(where).there = GlobalVar(nm$, dummy, , glb)
    
    If Typename(myVar(where).Value) = "lambda" Then
         
                        If glb Then
                        GlobalSub nm$ + "()", "CALL EXTERN " & CStr(myVar(where).there)
                    Else
                        GlobalSub HERE$ & "." & bstack.GroupName & nm$ + "()", "CALL EXTERN " & CStr(myVar(where).there)
                End If
    Set v(myVar(where).there) = myVar(where).Value
    ElseIf Typename(myVar(where).Value) = "Group" Then
    Dim aa As Object
    Set aa = myVar(where).Value
     UnFloatGroup bstack, nm$, myVar(where).there, aa
     Set aa = Nothing
     Else
     Set v(myVar(where).there) = myVar(where).Value
    End If
    
    Else
    myVar(where).there = GlobalVar(nm$, myVar(where).Value, , glb)
    End If
Next i
End If
End Sub

Sub CopyFromVar(bstack As basetask, v() As Variant)
Dim i As Long, aa As Group
For i = 0 To used - 1
    If IsObject(v(myVar(i).there)) Then
        If Typename(myVar(where).Value) = "Group" Then
       
             ' using copygroup from module1
            Set bstack.lastobj = Nothing
            CopyGroup v(myVar(i).there), bstack
            Set myVar(where).Value = bstack.lastobj
        Else
        Set myVar(i).Value = v(myVar(i).there)
        Set v(myVar(i).there) = Nothing
        End If
        
    Else
        myVar(i).Value = v(myVar(i).there)
    End If
    myVar(i).there = 0
Next i

End Sub

Sub FeedNonLocal(aName$, val As Variant)
Dim where As Long
If Not myHash.Find(aName$, where) Then
where = GetFree()
myHash.ItemCreator aName$, where
End If
If IsObject(val) Then
'' maybe here we can copy
'' check first the type
    Set myVar(where).Value = val
Else
    myVar(where).Value = val
End If
End Sub

Private Sub Class_Initialize()
Total = 10
ReDim myVar(Total) As Vars
End Sub
Private Function GetFree() As Long
    If used >= Total Then
        Total = Total * 2
        ReDim Preserve myVar(Total) As Vars
    End If
    GetFree = used
    used = used + 1
End Function

Private Sub Class_Terminate()
Dim i As Long
If used > 0 Then
For i = used - 1 To 0
If IsObject(myVar(i).Value) Then
Set myVar(i).Value = Nothing
End If
Next i
Erase myVar()
End If
End Sub
Public Sub CopyTo(A As lambda, v() As Variant)
' only for new
Dim b As lambda, VR As Variant, where As Long
Set A = New lambda
A.code$ = Me.code$
Dim i As Long
If used > 0 Then
For i = 0 To used - 1
myHash.ReadVar i, nm$, where
If myVar(where).there <> 0 Then
If Typename(v(myVar(where).there)) = "lambda" Then
    With v(myVar(where).there)
    .CopyTo b, v()
    End With
    Set VR = b
    A.FeedNonLocal nm$, VR
    Set VR = Nothing
    Set b = Nothing
Else
    A.FeedNonLocal nm$, v(myVar(where).there)
End If
Else
If Typename(myVar(where).Value) = "lambda" Then
    myVar(where).Value.CopyTo b, v()
    Set VR = b
    A.FeedNonLocal nm$, VR
    Set VR = Nothing
    Set b = Nothing
Else
    A.FeedNonLocal nm$, myVar(where).Value
End If
End If
Next i
End If
End Sub
