VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTestClass2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objEvTest As clsTestClass1
Attribute objEvTest.VB_VarHelpID = -1


Private Sub Class_Initialize()
    MsgBox "clsTestClass2 Class_Initialize()"
    Set objEvTest = New clsTestClass1
End Sub

Public Sub Raise(ByVal strTest As String)
    Call objEvTest.Raise(strTest)
End Sub

Private Sub Class_Terminate()
    MsgBox "clsTestClass2 Class_Terminate()"
End Sub


Private Sub objEvTest_TestingEvent(ByVal strTest As String)
    MsgBox "TestingEvent callback, strTest = " & strTest
End Sub
