VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mStiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' strings and numbers goes to Collection
Private mstack As New Collection
Private mMany As Long
Dim BeginFrom As Long
' interface for iteration
Public index As Long

' Generic is base 0 in Stack Item
' Here we have base 1
'Stiva.PushStr   string or reference
'Stiva.DataStr
'Stiva.PushVal
'Stiva.DataVal
'Stiva.PopType  (η κορυφή)
'Stiva.PopStr
'Stiva.PopVal
'Stiva.Flush
Public Function IsEmpty() As Boolean
IsEmpty = mstack.count = 0
End Function
Public Property Get StackItem(where As Long) As Variant
Dim aa As VarItem
On Error Resume Next
Set aa = mstack.item(where)
Select Case aa.Typename
Case "S"
StackItem = aa.ItemStr
Case "N", "?", "L"
StackItem = aa.ItemDouble
Case Else
Set StackItem = aa.ObjectRef
End Select
End Property
Public Sub MakeTopItemBack(where As Long)
'we pick an item from top to any other place
Dim aa As VarItem
On Error Resume Next
If where > 1 Then
Set aa = mstack.item(1)
mstack.Remove 1
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , , where - 1
End If
End If
End Sub
Public Sub MakeTopItem(where As Long)
'we pick an item and to the top
Dim aa As VarItem
On Error Resume Next
If where > 1 Then
Set aa = mstack.item(where)
mstack.Remove where
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
End If
End Sub
Public Sub Copy2TopItem(where As Long)
'we pick an item and to the top
Dim aa As VarItem, mm As Object, mn As mArray, work As Object
On Error Resume Next
Set work = mstack.item(where).ObjectRef
Select Case Typename(work)
Case "Group"
    Set aa = New VarItem
    aa.Typename = "*"
    CopyGroupObj work, mm
    Set aa.ObjectRef = mm
Case "lambda"
    Set aa = New VarItem
    aa.Typename = "*"
    ' use copylambdaAny from module 1
    CopyLambdaAny work, mm
    Set aa.ObjectRef = mm
Case "mArray"
    Set aa = New VarItem
    aa.Typename = "*"
    Set mn = New mArray
    work.CopyArray mn
    Set aa.ObjectRef = mn
Case Else
Set aa = mstack.item(where)
End Select
Set work = Nothing
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub
Public Function CopyMe() As mStiva
Set CopyMe = New mStiva
CopyMe.Copy2TopNItems2FromStiva Me.Total, Me
End Function
Public Function CopyMe2(ByVal many As Long) As mStiva
Set CopyMe2 = New mStiva
If many < 0 Then
many = -many
CopyMe2.Copy2BottomNItems2FromStiva many, Me
Else
CopyMe2.Copy2TopNItems2FromStiva many, Me
End If
End Function
Public Function MergeFromOther(that As Object) As mStiva
Dim mm As mStiva
Set mm = that
Set MergeFromOther = CopyMe()
MergeFromOther.MergeBottom mm
End Function
Public Sub Copy2TopNItems2FromStiva(many As Long, Source As mStiva)
'we pick an item and to the top
Dim aa As VarItem, mm As Object, mn As mArray, work As Object
On Error Resume Next
For many = many To 1 Step -1
Set work = Source.objref(many)
Select Case Typename(work)
Case "Group"
    Set aa = New VarItem
    aa.Typename = "*"
    CopyGroupObj work, mm
    Set aa.ObjectRef = mm
Case "lambda"
    Set aa = New VarItem
    aa.Typename = "*"
    ' use copylambdaAny from module 1
    CopyLambdaAny work, mm
    Set aa.ObjectRef = mm
Case "mArray"
    Set aa = New VarItem
    aa.Typename = "*"
    Set mn = New mArray
    work.CopyArray mn
    Set aa.ObjectRef = mn
Case Else
If Source.Total >= many Then

Set aa = Source.StackPickRef(many)
Else
Exit For
End If
End Select
Set work = Nothing
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
Next many
End Sub
Public Sub Copy2BottomNItems2FromStiva(many1 As Long, Source As mStiva)
' reverse order
Dim aa As VarItem, mm As Object, mn As mArray, work As Object, many As Long
On Error Resume Next
For many = Source.Total - many1 + 1 To Source.Total
Set work = Source.objref(many)
Select Case Typename(work)
Case "Group"
    Set aa = New VarItem
    aa.Typename = "*"
    CopyGroupObj work, mm
    Set aa.ObjectRef = mm
Case "lambda"
    Set aa = New VarItem
    aa.Typename = "*"
    ' use copylambdaAny from module 1
    CopyLambdaAny work, mm
    Set aa.ObjectRef = mm
Case "mArray"
    Set aa = New VarItem
    aa.Typename = "*"
    Set mn = New mArray
    work.CopyArray mn
    Set aa.ObjectRef = mn
Case Else
If Source.Total >= many Then

Set aa = Source.StackPickRef(many)
Else
Exit For
End If
End Select
Set work = Nothing
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
Next many
End Sub
Public Property Get StackPick(where As Long) As VarItem
Dim a As VarItem
dec
Set a = mstack.item(where)
mstack.Remove where

Set StackPick = a
End Property
Public Property Get StackPickRef(where As Long) As VarItem
Set StackPickRef = mstack.item(where)
End Property
Public Property Get StackItemType(where As Long) As String
Dim aa As VarItem
On Error Resume Next

Set aa = mstack.item(where)
StackItemType = aa.Typename
End Property
Public Property Get StackItemTypeObjectType(where As Long) As String
Dim aa As VarItem, bb As Object
On Error Resume Next

Set aa = mstack.item(where)
Select Case aa.Typename
Case "N", "S", "?", "L"
StackItemTypeObjectType = vbNullString
Case "*"
Set bb = aa.ObjectRef
If TypeOf bb Is mHandler Then
Set bb = bb.objref
CheckDeepAny bb
If bb.t1 = 3 Then
StackItemTypeObjectType = "*[" + Typename(bb) + "]"
ElseIf bb.t1 = 1 Then
StackItemTypeObjectType = "*[Inventory]"
Else
StackItemTypeObjectType = "*[Buffer]"
End If
Else
StackItemTypeObjectType = Typename(bb)
End If
Set bb = Nothing
Case Else
Set bb = aa.ObjectRef
StackItemTypeObjectType = Typename(bb)
Set bb = Nothing
End Select
End Property
Public Property Get StackItemTypeIsObject(where As Long) As Boolean
Dim aa As VarItem, bb As Object
On Error Resume Next

Set aa = mstack.item(where)
Select Case aa.Typename
Case "N", "S", "L", "?"
Case Else
StackItemTypeIsObject = True
End Select
End Property
Public Property Get Total() As Long
If mMany < 0 Then mMany = 0
Total = mMany
End Property
Public Property Get count() As Long
count = mMany
End Property
Private Sub Inc()
mMany = mMany + 1
End Sub
Private Sub dec()

mMany = mMany - 1
End Sub
Public Sub drop(ByVal items As Long)
Dim num As Long
On Error Resume Next
mMany = mMany - items
If mMany < 0 Then mMany = 0: Set mstack = New Collection: Exit Sub
For num = 1 To items
    mstack.Remove 1
Next num

End Sub
Public Sub Flush()

Dim num As Long
For num = 1 To mstack.count
dec
  mstack.Remove 1
 
Next num
mMany = 0
End Sub
Public Sub PushUndefine()
Dim aa As New VarItem
aa.Typename = "?"
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub

Public Sub PushVal(a As Variant)
Dim aa As New VarItem
aa.Typename = "N"
aa.ItemDouble = a
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub
Public Sub PushOptional()
Dim aa As New VarItem
aa.Typename = ">"
aa.ItemDouble = 0
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub
Public Sub PushValLong(a As Variant)
Dim aa As New VarItem
aa.Typename = "L"
aa.ItemVariant = a
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub
Public Sub PushObj(RHS As Object)
Dim aa As New VarItem
aa.Typename = "*"  'something
Set aa.ObjectRef = RHS
Set RHS = Nothing
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub
Public Sub PushObjVarItem(RHS As Object)
If mstack.count = 0 Then
mstack.Add RHS
Else
mstack.Add RHS, , 1
End If
Inc
End Sub
Public Property Get PopObj() As Object
On Error Resume Next
dec
Set PopObj = mstack.item(1).ObjectRef
mstack.Remove 1   '
End Property
Public Property Get objref(where As Long) As Object
On Error Resume Next
Set objref = mstack.item(where).ObjectRef
End Property
Public Sub PushStr(ByRef a As String)
Dim aa As New VarItem
aa.Typename = "S"
aa.ItemStr = a
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub
Public Function LookNull() As Boolean
On Error Resume Next
If mstack.count = 0 Then
LookNull = False
Else
LookNull = mstack.item(1).NoValue

End If
End Function
Public Function PopType() As String
On Error Resume Next
If mstack.count = 0 Then
PopType = vbNullString
Else
PopType = mstack.item(1).Typename
End If
End Function
Public Property Get LookTopVal() As Variant
On Error Resume Next
If mMany > 0 Then

LookTopVal = mstack.item(1).ItemDouble
End If
End Property
Public Property Get LookNVal(n As Long) As Variant
On Error Resume Next
LookNVal = mstack.item(n).ItemVariant
End Property

Public Property Get PopVal() As Variant
On Error Resume Next
dec
PopVal = mstack.item(1).ItemDouble
mstack.Remove 1
End Property
Public Property Get PopVariant() As Variant
On Error Resume Next
dec
PopVariant = mstack.item(1).ItemVariant
mstack.Remove 1
End Property
Public Property Get PopValLong() As Variant
On Error Resume Next
dec
PopValLong = CLng(mstack.item(1).ItemDouble)
mstack.Remove 1
End Property
Public Property Get PopStr() As String
On Error Resume Next
dec
PopStr = mstack.item(1).ItemStr
mstack.Remove 1

End Property
Public Sub DataValLong(a As Variant)
Dim aa As New VarItem
aa.Typename = "L"
aa.ItemVariant = a
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , , mstack.count
End If
Inc
End Sub
Public Sub DataVal(a As Variant)
Dim aa As New VarItem
aa.Typename = "N"
aa.ItemVariant = a
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , , mstack.count
End If
Inc
End Sub
Public Sub DataOptional()
Dim aa As New VarItem
aa.Typename = ">"
aa.ItemDouble = 0
    If mstack.count = 0 Then
        mstack.Add aa
    Else
        mstack.Add aa, , , mstack.count
    End If
Inc
End Sub
Public Sub DataStr(ByRef a As String)
Dim aa As New VarItem
aa.Typename = "S"
aa.ItemStr = a$


If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , , mstack.count
End If
Inc
End Sub
Public Sub DataStrUn(a As String)
Dim aa As New VarItem, i As Long, b As String
b = Split(a)(0)
aa.Typename = "S"
aa.ItemStr = a
If mstack.count = 0 Then
mstack.Add aa
Else
For i = 1 To Total
If Split(StackItem(i))(0) = b Then

mstack.Remove i
If mstack.count = 0 Then
mstack.Add aa
ElseIf i = 1 Then
mstack.Add aa, , 1
Else
mstack.Add aa, , , i - 1
End If
Exit Sub
End If
Next i
mstack.Add aa, , , mstack.count
End If
Inc
End Sub

Public Sub DataObj(RHS As Object)
Dim aa As VarItem
If RHS Is Nothing Then
    Set aa = New VarItem
    aa.Typename = "*"  'something
    Set aa.ObjectRef = RHS
ElseIf TypeOf RHS Is VarItem Then
    Set aa = RHS
    Set RHS = Nothing
Else
    Set aa = New VarItem
    aa.Typename = "*"  'something
    Set aa.ObjectRef = RHS
    Set RHS = Nothing
End If
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , , mstack.count
End If
Inc
End Sub
Public Sub DataObjVaritem(RHS As Object)
If mstack.count = 0 Then
mstack.Add RHS
Else
mstack.Add RHS, , , mstack.count
End If
Inc
End Sub
Public Sub Begin(from As Long)
BeginFrom = from
End Sub
Public Sub additem(ByRef a As String)
Dim aa As New VarItem
    aa.Typename = "S"
    aa.ItemStr = a$
    If mstack.count = 0 Then
        BeginFrom = 1
        mstack.Add aa
    ElseIf BeginFrom > mstack.count Then
        mstack.Add aa
    Else
        If BeginFrom = 1 Then
            mstack.Add aa, , 1
        Else
         mstack.Add aa, , , BeginFrom - 1
        End If
    End If
        BeginFrom = BeginFrom + 1

Inc
End Sub
Public Sub MergeTop(a As mStiva)
If Me Is a Then Exit Sub
Dim i As Long, aa As VarItem

If a.Total > 0 Then
Do
i = a.Total
Set aa = a.StackPick(i)
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1  'push so first of a is first on mstack
End If
Inc
Loop Until i <= 1
End If
End Sub
Public Sub MergeBottomCopyArray(a As mArray)
If a.count = 0 Then Exit Sub
Dim x1 As Long, vvl As Variant, aa As Object
Dim subgroup As Object, sub2 As Object, myArray As mArray, mySecondArray As mArray
 For x1 = 0 To a.count - 1
                                        
                    If a.MyIsObject(a.item(x1)) Then
                    Set vvl = a.item(x1)
                    If Typename(vvl) = "Group" Then
                        Set subgroup = vvl
                        CopyGroupObj subgroup, sub2
                        Set subgroup = Nothing
                        Set vvl = sub2
                        Set sub2 = Nothing
                    ElseIf Typename(vvl) = "mArray" Then
                        Set myArray = vvl
                        Set mySecondArray = New mArray
                        myArray.CopyArray mySecondArray
                        Set myArray = Nothing
                        Set vvl = mySecondArray
                        Set mySecondArray = Nothing
                    ElseIf Typename(vvl) = "VarItem" Then
                            DataOptional
                            GoTo here
                    End If
                    ' DO SOMETHING
                        Set aa = vvl
                        DataObj aa
                        Set vvl = Nothing
                        Set aa = Nothing
                    Else
                        If a.MyIsNumeric(a.item(x1)) Then
                        
                            DataVal a.item(x1)
                        Else
                            DataStr a.item(x1)
                        End If
                     End If
here:
                Next x1


End Sub
Public Sub MergeTopCopyArray(a As mArray)
If a.count = 0 Then Exit Sub
Dim x1 As Long, vvl As Variant, aa As Object
Dim subgroup As Object, sub2 As Object, myArray As mArray, mySecondArray As mArray
 For x1 = 0 To a.count - 1
                                        
                    If a.MyIsObject(a.item(x1)) Then
                    Set vvl = a.item(x1)
                    If Typename(vvl) = "Group" Then
                        Set subgroup = vvl
                        CopyGroupObj subgroup, sub2
                        Set subgroup = Nothing
                        Set vvl = sub2
                        Set sub2 = Nothing
                    ElseIf Typename(vvl) = "mArray" Then
                        Set myArray = vvl
                        Set mySecondArray = New mArray
                        myArray.CopyArray mySecondArray
                        Set myArray = Nothing
                        Set vvl = mySecondArray
                        Set mySecondArray = Nothing
                    End If
                    ' DO SOMETHING
                        Set aa = vvl
                        PushObj aa
                        Set vvl = Nothing
                        Set aa = Nothing
                    Else
                        If a.MyIsNumeric(a.item(x1)) Then
                            PushVal a.item(x1)
                        Else
                            PushStr a.item(x1)
                        End If
                     End If
    
                Next x1


End Sub
Function ExportArray(ByVal items As Long) As mArray

If items > count Then items = count
Set ExportArray = New mArray
If items = 0 Then Exit Function
Dim i As Long, aa As VarItem
With ExportArray
    .myarrbase = 0
    .PushDim (items)
    .PushEnd
    .Arr = True


For i = 1 To items
Set aa = mstack.item(i)
Select Case aa.Typename
Case "S"
.item(i - 1) = aa.ItemStr
Case "N", "?", "L"
.item(i - 1) = aa.ItemDouble
Case Else
Set .item(i - 1) = aa.ObjectRef
End Select
Next i
End With
drop items

End Function

Public Sub MergeBottom(a As mStiva)
If Me Is a Then Exit Sub
Dim i As Long, aa As VarItem
If a.Total > 0 Then
Do
Set aa = a.StackPick(1)
If mstack.count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , , mstack.count    'push so first of a is first on mstack
End If
Inc
Loop Until a.Total < 1
End If
End Sub
Private Sub Class_Terminate()
Flush
End Sub
Sub CopyGroupObj(myobject1 As Object, myobject2 As Object)
 Dim ps As New mStiva, v As Long, s$, frm$, vvl As Variant, x1 As Long, ss$
Dim subgroup As Object, sub2 As Object, myArray As mArray, mySecondArray As mArray
 Dim mygroup2 As New Group
  Dim myobject As Group
  
 Set myobject = myobject1
 If myobject.IamApointer Then Set myobject2 = myobject1: Exit Sub
             With myobject
                If .IamFloatGroup Then
         
                 .PeekItem 1, vvl
                 mygroup2.BeginFloat vvl + 2
                 For x1 = 0 To vvl * 2 + 3
                    .PeekItem x1, vvl
                    If Typename(vvl) = "Group" Then
                    Set subgroup = vvl
                    CopyGroupObj subgroup, sub2
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
                    mygroup2.PokeItem x1, vvl
     
                Next x1
                End If
        
                mygroup2.HasStrValue = .HasStrValue
                mygroup2.HasValue = .HasValue
                mygroup2.HasSet = .HasSet
                mygroup2.HasParameters = .HasParameters

             Set mygroup2.SuperClassList = .SuperClassList
                Set mygroup2.Events = .Events
                mygroup2.highpriorityoper = .highpriorityoper
                mygroup2.HasUnary = .HasUnary
            End With
            Set myobject2 = mygroup2
End Sub
Function Copy() As mStiva
Dim i As Long
Set Copy = New mStiva
For i = Total To 1 Step -1
'Debug.Print Me.StackItem(i)
Copy.PushStr Me.StackItem(i)
Next i
End Function
Property Get Value() As Variant
Dim aa As VarItem
If index >= 0 And index < mMany Then
    Set aa = StackPickRef(index + 1)
    If aa.ObjectRef Is Nothing Then
        If aa.Typename = "S" Then
            Value = aa.ItemStr
        Else
            Value = aa.ItemDouble
        End If
    Else
        Set Value = aa.ObjectRef
    End If
End If
End Property
Sub SplitStackTop(fromstiva As Object, item As Long)
Dim aa As VarItem, i As Long
Dim thisstiva As mStiva
Set thisstiva = fromstiva.objref
If item > thisstiva.count Then item = thisstiva.count
For i = 1 To item
    Set aa = thisstiva.StackPick(1)
        If mstack.count = 0 Then
        mstack.Add aa
        Else
        mstack.Add aa, , , mstack.count
        End If
        Inc
Next i
End Sub
Sub SplitStackBottom(fromstiva As Object, item As Long)
Dim aa As VarItem, i As Long
Dim thisstiva As mStiva
Set thisstiva = fromstiva.objref
If item < 1 Then item = 1
For i = item To thisstiva.count
    Set aa = thisstiva.StackPick(item)
        If mstack.count = 0 Then
        mstack.Add aa
        Else
        mstack.Add aa, , , mstack.count
        End If
        Inc
Next i
End Sub

