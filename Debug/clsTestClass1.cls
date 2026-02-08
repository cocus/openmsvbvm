VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTestClass1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event TestingEvent(ByVal strTest As String)


Private Sub Class_Initialize()
    MsgBox "clsTestClass1 Class_Initialize()"
End Sub

Public Sub Raise(ByVal strTest As String)
    RaiseEvent TestingEvent(strTest)
End Sub

Private Sub Class_Terminate()
    MsgBox "clsTestClass1 Class_Terminate()"
End Sub

