VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mmname As String
Dim dnum As Long, ndnum As Long
Dim limit(10) As Long  ' max 9 dimensions
Dim dtable(10) As Long
Dim nlimit(10) As Long  ' max 9 dimensions

Dim actualdata() As Variant
Dim maxitems As Long
Dim maxitems1 As Long
Dim nmaxitems As Long
Dim nmaxitems1 As Long
Public myarrbase As Long
Public common As Boolean
Public IHaveClass As Boolean
Public IHaveGui As Boolean
Public GroupRef As Object   ' get a floating Group, maybe with groups inside....as floating group
Public CodeName As String
Private oldmaxitems As Long
Public Arr As Boolean
Public RevOrder As Boolean
Private Type SAFEArrayPart
   cDims                        As Integer
   fFeatures                    As Integer
   cbElements                   As Long
   cLocks                       As Long
   pvData                       As Long
   End Type
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal addr As Long, retval As Long)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Long)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal addr As Long, retval As Integer)

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
  Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" _
      (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
        Private Declare Sub CopyPart Lib "KERNEL32" Alias "RtlMoveMemory" _
      (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SafeArrayGetDim Lib "OleAut32.dll" (ByVal pSA As Long) As Long
Private Declare Function SafeArrayLock Lib "OleAut32.dll" (ByVal pSA As Long) As Long
Private Declare Function SafeArrayUnlock Lib "OleAut32.dll" (ByVal pSA As Long) As Long
  
Private Declare Function SafeArrayAccessData Lib "OleAut32.dll" (ByVal pSA As Long, ppVdata As Long) As Long
Private Declare Function SafeArrayUnaccessData Lib "OleAut32.dll" (ByVal pSA As Long) As Long
Private Declare Function SafeArrayGetElemsize Lib "OleAut32.dll" (ByVal pSA As Long) As Long
Private Declare Sub SafeArrayCopyData Lib "OleAut32.dll" (ByVal psaSource As Any, ByVal psaTarget As Any)
Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLBound As Long
End Type
Public index As Long
Public Sub Compute(p As Double, ss$)
Dim v As Long
On Error Resume Next
v = 0
Do While v <= Me.UpperMonoLimit
If MyIsNumeric(actualdata(v)) Then
                    Select Case ss$
                                Case "=", "g"
                                actualdata(v) = p
                                Case "+"
                                actualdata(v) = p + actualdata(v)
                                Case "*"
                                 actualdata(v) = p * actualdata(v)
                                Case "-"
                                actualdata(v) = actualdata(v) - p
                                Case "/"
                                If p = 0 Then Exit Do
                                 actualdata(v) = actualdata(v) / p
                                  Case "~", "!!"
                                 actualdata(v) = -1 - (actualdata(v) <> 0)
                                 Case "++"
                                actualdata(v) = 1 + actualdata(v)
                                Case "--"
                                actualdata(v) = actualdata(v) - 1
                                  Case "-!"
                                actualdata(v) = -actualdata(v)
                                  
                                 End Select
                                 actualdata(v) = MyRound(actualdata(v), 13)
                            
End If
                           v = v + 1
                           Err.Clear
Loop

End Sub
Public Sub Compute2(p As Double, ss$)
Dim v As Long
On Error Resume Next
v = 0
Do While v <= Me.UpperMonoLimit
If MyIsNumeric(actualdata(v)) Then
                    Select Case ss$
                                Case "="
                                actualdata(v) = p
                                Case "+="
                                actualdata(v) = p + actualdata(v)
                                Case "*="
                                 actualdata(v) = p * actualdata(v)
                                Case "-="
                                actualdata(v) = actualdata(v) - p
                                Case "/="
                                If p = 0 Then Exit Do
                                 actualdata(v) = actualdata(v) / p
                                 
                                 End Select
                                 actualdata(v) = MyRound(actualdata(v), 13)
                            
End If
                           v = v + 1
                           Err.Clear
Loop

End Sub
Public Sub Compute3(ss$)
Dim v As Long
On Error Resume Next
v = 0
Do While v <= Me.UpperMonoLimit
If MyIsNumeric(actualdata(v)) Then
                    Select Case ss$
                                  Case "~"
                                 actualdata(v) = -1 - (actualdata(v) <> 0)
                                 Case "++"
                                actualdata(v) = 1 + actualdata(v)
                                Case "--"
                                actualdata(v) = actualdata(v) - 1
                                  Case "-!"
                                actualdata(v) = -actualdata(v)
                                  
                                 End Select
End If
                           v = v + 1
                           Err.Clear
Loop

End Sub
Property Get Value() As Variant
If index >= 0 And index <= maxitems1 Then

If IsObject(actualdata(index)) Then
    If Typename(actualdata(index)) = "PropReference" Then
        Value = actualdata(index)
    Else
        Set Value = actualdata(index)
    End If
Else
Value = actualdata(index)
End If
End If
End Property
Public Sub PushProp(a As Variant)
Set GroupRef = a
Arr = False
End Sub
Public Sub CopySerialize(MMM() As Variant)
' may we get an exception here
Dim m_base As Long, first12 As Long
Dim m_base_0 As Long, first12_0 As Long, r1 As Long, r2 As Long

GetMem4 VarPtrArray(MMM()), m_base
GetMem4 VarPtrArray(actualdata()), m_base_0
r1 = SafeArrayLock(m_base)
r2 = SafeArrayLock(m_base_0)
r1 = SafeArrayAccessData(ByVal m_base, first12)
r2 = SafeArrayAccessData(ByVal m_base_0, first12_0)

'GetMem4 m_base + 12, first1
'GetMem4 m_base_0 + 12, first1_0
PutMem4 m_base + 12, first12_0
PutMem4 m_base_0 + 12, first12
r2 = SafeArrayUnaccessData(ByVal m_base_0)
r1 = SafeArrayUnaccessData(ByVal m_base)
r2 = SafeArrayUnlock(m_base_0)
r1 = SafeArrayUnlock(m_base)
End Sub
Public Sub CopyBackSerialize(MMM() As Variant)
' may we get an exception here
Dim m_base As Long, first12 As Long
Dim m_base_0 As Long, first12_0 As Long, r1 As Long, r2 As Long
If Not RevOrder Then Err.Raise vbObjectError + 12345, "CopyBack", "Failed. is not vb6"
Dim backup()
backup() = actualdata()
GetMem4 VarPtrArray(MMM()), m_base
GetMem4 VarPtrArray(backup()), m_base_0
r1 = SafeArrayLock(m_base)
r2 = SafeArrayLock(m_base_0)
r1 = SafeArrayAccessData(ByVal m_base, first12)
r2 = SafeArrayAccessData(ByVal m_base_0, first12_0)

'GetMem4 m_base + 12, first1
'GetMem4 m_base_0 + 12, first1_0
PutMem4 m_base + 12, first12_0
PutMem4 m_base_0 + 12, first12
r2 = SafeArrayUnaccessData(ByVal m_base_0)
r1 = SafeArrayUnaccessData(ByVal m_base)
r2 = SafeArrayUnlock(m_base_0)
r1 = SafeArrayUnlock(m_base)
End Sub
Public Sub Car(m As Object)
Dim k As New mArray
Dim v As Variant, I As Long, obj As Object, obj2 As Object, spare As lambda
If Count > 0 Then
    k.PushDim 1
    k.PushEnd
    k.SerialItem actualdata(0), 0, 1
    If IsObject(Me.item(0)) Then
    Select Case Typename(Me.item(0))
    Case "Group"
    Set obj = Me.item(0)
    CopyGroup obj, obj2
    Set k.item(0) = obj2
    Set obj2 = Nothing
    Case "lambda"
    Dim actualdata1()
    With Me.item(0)
    .CopyTo spare, actualdata1()
    End With
    Set k.item(0) = spare
    Set spare = Nothing
    Case Else
    If TypeOf Me.item(0) Is mHandler Then
        If Me.item(0).t1 = 3 Then
        'Set m = New mHandler
        'm.t1 = 3
        Set m = Me.item(0)
       ' If Not CheckIsmArray(m, var()) Then
        '    Set m = New mHandler
         '   m.t1 = 3
          '  Set m.objref = Me.item(0)
        'End If
        Set k = Nothing
        Exit Sub
        End If
     ElseIf TypeOf Me.item(0) Is mArray Then
             Set m = New mHandler
        m.t1 = 3
        Set m.objref = Me.item(0)
        Set k = Nothing
        Exit Sub
    End If
    Set k.item(0) = Me.item(0)
    
    End Select
    Else
    k.item(0) = Me.item(0)
    End If
End If
k.Arr = True
Set m = k
Set k = Nothing
End Sub
Public Sub Cdr(m As Object)
Dim k As New mArray
Dim v As Variant, I As Long, obj As Object, obj2 As Object, spare As lambda
If dnum > 0 Then
k.StartResize: k.PushDim Me.Count - 1: k.PushEnd
Else
k.Arr = True
Set k.GroupRef = Me.GroupRef
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
Set m = k
Exit Sub
End If
k.myarrbase = myarrbase
k.RevOrder = False
For I = 1 To Me.UpperMonoLimit
If IsObject(Me.item(I)) Then
Select Case Typename(Me.item(I))
Case "Group"
Set obj = Me.item(I)
CopyGroup obj, obj2
Set k.item(I - 1) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata1()
With Me.item(I)
.CopyTo spare, actualdata1()
End With
Set k.item(I - 1) = spare
Set spare = Nothing
Case Else
Set k.item(I - 1) = Me.item(I)
End Select
Else
k.item(I - 1) = Me.item(I)
End If
Next I
'' new
Set k.GroupRef = Me.GroupRef
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
Set m = k

End Sub
Public Sub processAppend(m As Object)
Dim pppp As New mArray, pppp1 As mArray
Set pppp1 = m
If dnum > 0 Then
    Me.CopyArray pppp
    If Not pppp1.IsEmpty Then pppp1.AppendArray pppp
ElseIf Not pppp1.IsEmpty Then
    Set pppp = pppp1
End If
pppp.Arr = True
Set m = pppp
Set pppp = Nothing
Set pppp1 = Nothing
End Sub

Public Sub AppendArray(k As mArray)
'append this to k
If k.Count = 0 Then
Me.CopyArray k
Exit Sub
End If
Dim v As Variant, I As Long, obj As Object, obj2 As Object, spare As lambda
Dim fromhere As Long
fromhere = k.UpperMonoLimit + 1
k.SerialItem v, Me.Count + k.Count, 10
k.myarrbase = myarrbase
k.RevOrder = RevOrder
If Right$(k.arrname, 2) = "%(" Then
For I = 0 To Me.UpperMonoLimit
If IsObject(Me.item(I)) Then
Select Case Typename(Me.item(I))
Case "Group"
Set obj = Me.item(I)
CopyGroup obj, obj2
Set k.item(I + fromhere) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata()
With Me.item(I)
.CopyTo spare, actualdata()
End With
Set k.item(I + fromhere) = spare
Set spare = Nothing
Case Else
Set k.item(I + fromhere) = Me.item(I)
End Select
Else
k.item(I + fromhere) = Int(Me.item(I))
End If
Next I
Else
''''''''''****************************
For I = 0 To Me.UpperMonoLimit

If IsObject(Me.item(I)) Then
Select Case Typename(Me.item(I))
Case "Group"
Set obj = Me.item(I)
CopyGroup obj, obj2
Set k.item(I + fromhere) = obj2
Set obj2 = Nothing
Case Else
Set k.item(I + fromhere) = Me.item(I)
End Select
Else
k.item(I + fromhere) = Me.item(I)
End If
Next I
End If
'' new
Set k.GroupRef = Me.GroupRef
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
End Sub



Public Sub CopyArray(k As mArray)
'copy this to k
Dim v As Variant, I As Long, obj As Object, obj2 As Object, spare As lambda
k.SerialItem v, Me.UpperMonoLimit, 8
k.myarrbase = myarrbase
k.RevOrder = False
k.StartResize
For I = 0 To dnum - 1
k.PushDim limit(I)
Next I

k.PushEnd
k.RevOrder = RevOrder
If Right$(k.arrname, 2) = "%(" Then
For I = 0 To Me.UpperMonoLimit

If IsObject(Me.item(I)) Then
Select Case Typename(Me.item(I))
Case "Group"
Set obj = Me.item(I)
CopyGroup obj, obj2
Set k.item(I) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata()
With Me.item(I)
.CopyTo spare, actualdata()
End With
Set k.item(I) = spare
Set spare = Nothing

Case Else
Set k.item(I) = Me.item(I)
End Select
Else
k.item(I) = Int(Me.item(I))
End If
Next I
Else
''''''''''****************************
For I = 0 To Me.UpperMonoLimit

If IsObject(Me.item(I)) Then
Select Case Typename(Me.item(I))
Case "Group"
Set obj = Me.item(I)
CopyGroup obj, obj2
Set k.item(I) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata1()
With Me.item(I)
.CopyTo spare, actualdata1()
End With
Set k.item(I) = spare
Set spare = Nothing
Case Else
Set k.item(I) = Me.item(I)
End Select
Else
k.item(I) = Me.item(I)
End If
Next I
End If
'' new
Set k.GroupRef = Me.GroupRef
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
End Sub

Public Function UpperMonoLimit()
UpperMonoLimit = maxitems1
End Function
Public Property Get Count()
If dnum = 0 Then
Count = 0
Else
Count = maxitems1 + 1 ' sa
End If
End Property
Public Property Let arrname(aName As String)
mmname = aName
End Property
Public Property Get arrname() As String
arrname = mmname
End Property
Public Sub PushDim(nDim As Long)
If ndnum < 10 Then
nlimit(ndnum) = nDim
nmaxitems = nmaxitems * nDim
ndnum = ndnum + 1
End If
nmaxitems1 = nmaxitems - 1
End Sub
Public Sub PushEnd()
On Error GoTo there11
If ndnum = 0 Then
dnum = 0
maxitems = 1
maxitems1 = 0
ReDim actualdata(maxitems1) As Variant
Else
Dim I As Long
For I = 0 To 9: limit(I) = nlimit(I): Next I
    dnum = ndnum
    maxitems = nmaxitems
    maxitems1 = nmaxitems1

If nmaxitems1 < 0 Then
    ReDim actualdata(1) As Variant
Else
    ReDim Preserve actualdata(maxitems1) As Variant
End If
Dim mx As Long
mx = maxitems
I = 0
If limit(I) > 0 Then
If RevOrder Then
Dim revlimit()
ReDim revlimit(UBound(limit()))
    For I = 0 To dnum - 1
    revlimit(I) = limit(dnum - I - 1)
    Next I
    For I = 0 To dnum - 1
    limit(I) = revlimit(I)
    Next I
    
End If
For I = 0 To dnum - 1
    dtable(I) = mx / limit(I)
    mx = dtable(I)
Next I
Else
GoTo there11
End If

End If
Exit Sub
there11:
dnum = 0
maxitems = 1
maxitems1 = 0
ReDim actualdata(maxitems)
End Sub
Public Sub StartResize()
oldmaxitems = maxitems1
ndnum = 0 ' no actions
nmaxitems = 1
nmaxitems1 = 0
End Sub
Public Function SerialItem(item As Variant, Cursor As Long, command As Long) As Boolean
If command = 1 Then
If dnum = 0 Then Exit Function
If Cursor < 0 Or Cursor >= maxitems Then Exit Function
If VarType(item) = vbObject Then
Set actualdata(Cursor) = item
Else
actualdata(Cursor) = item
End If
ElseIf command = 2 Then
If dnum = 0 Then Exit Function
If Cursor < 0 Or Cursor >= maxitems Then Exit Function
item = actualdata(Cursor)
ElseIf command = 3 Then
    Fill item
ElseIf command = 4 Then
    Cursor = maxitems
ElseIf command = 5 Then
Cursor = dnum
item = dnum
ElseIf command = 6 Then
If Cursor > 9 Or Cursor < 0 Then
item = 0
Else
If RevOrder Then
item = limit(dnum - Cursor - 1)
Else
item = limit(Cursor)
End If
End If
ElseIf command = 7 Then  'erase all data
Dim aa As Variant
Fill aa
ElseIf command = 8 Then
If dnum = 1 Then
If Cursor = -1 Then
dnum = 0 ' no actions
maxitems = 1
maxitems1 = 0
ndnum = 0 ' no actions
nmaxitems = 1
nmaxitems1 = 0
ReDim actualdata(0)
CodeName = GetAnyName$
Arr = True
limit(0) = 0
Else
    ReDim Preserve actualdata(Cursor) As Variant
    
    maxitems = Cursor
End If
    limit(0) = maxitems

End If
ElseIf command = 9 Then
If dnum = 1 Then
    ReDim Preserve actualdata(Cursor) As Variant
    maxitems = Cursor - 1
    limit(0) = maxitems
End If
maxitems1 = maxitems - 1
ElseIf command = 10 Then
   
    ReDim Preserve actualdata(Cursor) As Variant
    maxitems = Cursor
    limit(0) = maxitems - 1
dnum = 1
maxitems1 = maxitems - 1
End If
SerialItem = True

End Function

Public Function PushOffset(curitem As Long, curdim As Long, nDim As Long) As Boolean
If curdim >= dnum Then
' error...
PushOffset = False
Else

If RevOrder Then
If nDim >= limit(dnum - curdim - 1) - myarrbase Then Exit Function
curitem = curitem + dtable(dnum - curdim - 1) * (nDim + myarrbase)
Else
If nDim >= limit(curdim) - myarrbase Then Exit Function
curitem = curitem + dtable(curdim) * (nDim + myarrbase)
End If
PushOffset = True
End If
End Function
Public Property Get itemPtr(curitem As Long) As Long
itemPtr = VarPtr(actualdata(curitem))
End Property
Public Property Get item(curitem As Long) As Variant
On Error GoTo there1
If Arr Then
If maxitems1 >= curitem Then
If IsObject(actualdata(curitem)) Then
If Typename(actualdata(curitem)) = "PropReference" Then
item = actualdata(curitem)
Else
Set item = actualdata(curitem)
End If
Else

item = actualdata(curitem)
End If
Exit Property
Else
''Stop
End If
Else
If curitem = -2 Then
Set item = GroupRef
ElseIf curitem < 0 Then
'If TypeOf GroupRef Is Group Then Set item = GroupRef: Exit Property
If GroupRef.objref.Done Then
If GroupRef.objref.IsObj Then
Set item = GroupRef.objref.ValueObj
Else
item = GroupRef.objref.Value
End If

End If
End If
End If
there1:

End Property
Public Property Get itemnumeric(curitem As Long) As Double
On Error GoTo there1
If Arr Then
If maxitems1 >= curitem Then
If IsObject(actualdata(curitem)) Then
If Typename(actualdata(curitem)) = "PropReference" Then
itemnumeric = CDbl(actualdata(curitem))
Else
itemnumeric = 0
End If
Else

itemnumeric = CDbl(actualdata(curitem))
End If
Exit Property
Else
''Stop
End If
Else
If curitem = -2 Then
ElseIf curitem < 0 Then
If GroupRef.objref.Done Then
If GroupRef.objref.IsObj Then
itemnumeric = 0
Else
itemnumeric = GroupRef.objref.Value
End If
End If
End If
End If
there1:

End Property

Public Property Let item(curitem As Long, item As Variant)
On Error Resume Next
If curitem < 0 Then

ElseIf Not IsObject(actualdata(curitem)) Then
actualdata(curitem) = item
Else
If Not IsObject(item) Then
Set actualdata(curitem) = Nothing
Else
actualdata(curitem).Value = item
End If
End If

End Property
Public Property Set item(curitem As Long, item As Variant)
On Error Resume Next
If curitem < 0 Then
If Not GroupRef Is Nothing Then
If TypeOf GroupRef Is mHandler Then
If GroupRef.t1 = 1 Then
If GroupRef.objref.Done Then
Set GroupRef.objref.ValueObj = item
Exit Property
End If
End If
End If
End If
MyEr "Wrong setting", "Λάθος όρισμα"
Else
Set actualdata(curitem) = item
End If
End Property

Private Sub Fill(item As Variant)
Dim aa As Object
If dnum > 0 Then
Dim I As Long
If IHaveClass Then
For I = 0 To maxitems - 1
CopyGroup GroupRef, aa
Set actualdata(I) = aa
Next I
Else
If Typename(item) = "Empty" Then
ReDim actualdata(maxitems)
''maxitems = 0

Else
For I = 0 To maxitems - 1
actualdata(I) = item
Next I
End If
End If
End If
End Sub
Public Sub FillLambda(bstack As basetask)
Dim obj As Variant, I As Long
Set obj = bstack.lastobj
    Set bstack.lastobj = Nothing
For I = 0 To maxitems - 1
 CopyLambda obj, bstack
Set actualdata(I) = bstack.lastobj
Set bstack.lastobj = Nothing
Next I
Set obj = Nothing
End Sub
Sub MakeitObject2(var As Variant)
Dim aa As Object
Set aa = New Group
Set var = aa
End Sub
Sub CopyGroup(myobject As Object, myobject2 As Object)
'Special Copy
 Dim ps As New mStiva, v As Long, s$, frm$, vvl As Variant, x1 As Long, ss$
Dim subgroup As Object, sub2 As Object, myArray As mArray, mySecondArray As mArray
 Set myobject2 = New Group
 
             With myobject
                If .IamFloatGroup Then
         
                 .PeekItem 1, vvl
                 myobject2.BeginFloat vvl + 2
                 For x1 = 0 To vvl * 2 + 2
                    .PeekItem x1, vvl
                    If Typename(vvl) = "Group" Then
                    Set subgroup = vvl
                    CopyGroup subgroup, sub2
                    Set subgroup = Nothing
                    Set vvl = sub2
                    Set sub2 = Nothing
                    ElseIf Typename(vvl) = "mArray" Then
                    If Not vvl.common Then
                        Set myArray = vvl
                        Set mySecondArray = New mArray
                        myArray.CopyArray mySecondArray
                        Set myArray = Nothing
                        Set vvl = mySecondArray
                        Set mySecondArray = Nothing
                    End If
                    End If
                    myobject2.PokeItem x1, vvl
     
                Next x1
                myobject2.HasStrValue = myobject.HasStrValue
                myobject2.HasValue = myobject.HasValue
                myobject2.HasSet = myobject.HasSet
                myobject2.HasParameters = myobject.HasParameters
                Set myobject2.LinkRef = myobject
                End If
            End With
         
End Sub
Property Get IsEmpty() As Variant
IsEmpty = dnum = 0
End Property
Private Sub Class_Initialize()
dnum = 0 ' no actions
maxitems = 1
maxitems1 = 0
ndnum = 0 ' no actions
nmaxitems = 1
nmaxitems1 = 0
ReDim actualdata(0)
CodeName = GetAnyName$
Arr = True
myarrbase = -ArrBase
End Sub

Private Sub Class_Terminate()
Set GroupRef = Nothing
Erase actualdata()
End Sub
Private Function GetAnyName$()
Dim ss$, j As Long
For j = 1 To 4
ss$ = ss$ & Chr(65 + Int((23 * Rnd) + 1))
Next j
GetAnyName$ = ChrW(&H1FFF) + ss$
End Function
Function MyIsNumeric(v As Variant) As Boolean
Dim n As Integer
GetMem2 VarPtr(v), n
If n < 2 Then Exit Function
MyIsNumeric = n < 8
End Function
Function MyIsObject(v As Variant) As Boolean
Dim n As Integer
GetMem2 VarPtr(v), n
MyIsObject = n = 9
End Function
