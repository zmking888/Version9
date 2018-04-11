VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VarItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mTypeName As String
Private mItem As Variant
Private slot As Integer
Private IamNull As Boolean
 Property Let Typename(ByRef RHS As String)
mTypeName = RHS
End Property

 Property Get Typename() As String
Typename = mTypeName
End Property

 Property Let ItemStr(ByRef RHS As String)
 IamNull = False
 slot = 4
mItem = CVar(RHS)
End Property

 Property Get ItemStr() As String
If slot = 4 Then ItemStr = mItem
End Property
 Property Let ItemDouble(ByVal RHS As Variant)
 IamNull = False
 slot = 2
mItem = RHS
End Property

Property Get ItemDouble() As Variant
If slot = 2 Then ItemDouble = mItem
End Property
 Property Let ItemVariant(RHS As Variant)
 slot = 2
 IamNull = False
mItem = RHS
End Property
 Property Get ItemVariant() As Variant
 If slot = 2 Then ItemVariant = mItem
End Property

Private Sub Class_Initialize()
IamNull = True
End Sub
Property Get NoValue() As Boolean
NoValue = IamNull
End Property

Public Property Set ObjectRef(RHS As Object)
IamNull = RHS Is Nothing
slot = 1
    Set mItem = RHS
End Property
Public Property Get ObjectRef() As Object
If slot = 1 Then Set ObjectRef = mItem
End Property

Private Sub Class_Terminate()
If slot = 1 Then Set mItem = Nothing
End Sub
