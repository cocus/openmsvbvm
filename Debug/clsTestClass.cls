VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTestClass3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private c_bTesting As Boolean

Private c_sString As String

'Public puto1 As Long
'Public puto2 As Long

'Private c_bTesting2 As Boolean

'Private c_bTesting22 As Boolean
'Private c_bTesting222 As Boolean

Private Sub Class_Initialize()
    MsgBox "Class_Initialize()"
    c_bTesting = True
End Sub

Public Sub SetTrue()
    c_bTesting = True
End Sub

Public Sub SetFalse()
    c_bTesting = False
End Sub

Public Sub SetString(ByVal sString As String)
    c_sString = sString
End Sub

Public Function MsgboxString(ByVal bYesNo As Boolean) As Boolean
    MsgBox c_sString & " " & CStr(c_bTesting)
    MsgboxString = True
End Function

Public Function Exported2() As Boolean
    MsgBox "Exported2Called!!! " & c_bTesting & " " & Hex(VarPtr(Me))
    Exported2 = False
End Function

Private Sub Class_Terminate()
    MsgBox "Class_Terminate()"
End Sub
