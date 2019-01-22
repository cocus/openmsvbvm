Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long


Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Type tTest
    sString         As String
    'sFixedString    As String * 30
    lLong           As Long
    iInteger        As Integer
End Type

Sub Main()
    'Call InitCommonControls
    
    Call OnErrorGotoTest
    
    'Call OnErrorResumeNextTest
    
    'Call FPExceptionTest
    
    'Call DllStringCallTest
    
    'Call MsgboxReturnCodeTest
    
    'Call StrCompTest
    
    'Dim s As String
    's = Replace$("This is a test", "test", "sample")
    'MsgBox s
    
    'Call ObjectCreateTest
    
    'Call DateTest

    'Call InputBoxTest
    
    'Call ToStringConversionTest
    
    'Call ObjectAppInteractionTest
    
    'Call FileManipulationTest
    
    'Call StringCopyTest
    
    'Call ClassCreationTest
    
    'Call ArraysTest
    
    'Call VBAInteractionTest
    
    'Call llmul
    
    Debug.Print "ASD"
    
    MsgBox "END"
End Sub

Private Sub llmul()
    Dim a As Currency
    Dim b As Currency
    Dim c As Currency
    a = -1.3333@
    b = 1.2@
    c = a * b
    'MsgBox c
End Sub
Private Sub ArraysTest()
    'Dim strArray() As String
    'ReDim strArray(36)
    'strArray(0) = "One"
    'strArray(1) = "Two"
    'strArray(2) = "Three"
    'ReDim Preserve strArray(69)
    'strArray(3) = "Four"
    'MsgBox strArray(3)
    'MsgBox "UBound() = " & UBound(strArray) & " - LBound() = " & LBound(strArray)
    
    
    Dim tUDTArray(30) As tTest
    'ReDim tUDTArray(30)
    'tUDTArray(0).sFixedString = "HAHAHA"
    With tUDTArray(0)
        .lLong = &HC0C0C0C0
        .sString = "string"
    End With
    
    With tUDTArray(0)
        MsgBox "tUDTArray(0)" & vbCrLf & _
                vbTab & ".sString = '" & .sString & "'" & vbCrLf & _
                vbTab & ".lLong = " & Hex(.lLong)
    End With
End Sub

Private Sub VBAInteractionTest()
    'VBA.AppActivate "Debug", 300
    'VBA.Beep
    Dim s(4) As Variant
    s(0) = "Zero"
    s(1) = "One"
    s(2) = "Two"
    s(3) = "Three"
    MsgBox VBA.Choose(2147483600, "Zero", "One", "Two", "Three")
    'MsgBox "Command() = " & Command() & vbCrLf & "Command$() = " & Command$()
    'SaveSetting "ItsMeMario", "Section", "Key", "Setting"
    'MsgBox "GetSetting = " & GetSetting("ItsMeMario", "Section", "Key") ', "ThisIsADefaultValue")
    'DeleteSetting "ItsMeMario", "Section", "Key"
    'MsgBox Environ("TMP") '("TMP")
    'MsgBox IIf(0, Environ("TMP"), Environ("PATH"))
    'MsgBox Partition(10, 0, 100, 5)
    'SendKeys "A", 333
    'MsgBox Switch(False, "should not happen 1", False, "should not happen 2", True, "Yes!!")
End Sub

Private Sub OnErrorGotoTest()
    On Error GoTo fnErrHandler
    'On Error Resume Next
    MsgBox "This should happen"
    MsgBox 1 / 0
fnNonErrHandler:
    MsgBox "This should not be visible"
fnErrHandler:
    MsgBox "OnErrorTest() end"
End Sub

Private Sub OnErrorResumeNextTest()
    'On Error Resume Next
    MsgBox "This should happen"
    MsgBox 1 / 0
fnNonErrHandler:
    MsgBox "This should not be visible"
fnErrHandler:
    MsgBox "OnErrorTest() end"
End Sub

Private Sub MsgboxReturnCodeTest()
    Select Case MsgBox("Just a messagebox and the command is " & Command$, vbInformation + vbYesNoCancel, "Just the title")
        Case vbYes: MsgBox UCase$("Yes"): Shell "calc"
        Case vbNo: MsgBox LCase$("No")
        Case vbCancel: MsgBox "Cancel"
    End Select
End Sub

Private Sub StringCaseTest()
    MsgBox LCase$("LCASE") & " " & UCase$("ucase")
End Sub

Private Sub StrCompTest()
    Dim l As Long
    l = StrComp("lcase", "XXXCASE") ', vbTextCompare)
    MsgBox "strcomp = " & l
End Sub

Private Sub DllStringCallTest()
    MessageBox 0, "Testing", "Title", vbOK
End Sub

Private Sub FPExceptionTest()
    'MsgBox 2147483647 + 1
    MsgBox 1 / 0
End Sub

Private Sub DateTest()
    Dim t As Date
    t = Now
    MsgBox "asd " & t
End Sub

Private Sub ClassCreationTest()
    Dim j As clsTestClass3
    Set j = New clsTestClass3
    'j.Pelotudo True
    'MsgBox "delay?"
    'j.Exported2
    'j.SetTrue
    'j.Exported2
    'j.SetFalse
    'j.Exported2
    
    j.SetString "TESTTTTTTTT!!!!!"
    j.MsgboxString True
    
    Set j = Nothing
    
End Sub

Private Sub StringCopyTest()
    Dim s As String
    Dim s2 As String
    s = "This is S"
    s2 = s
    MsgBox s2
End Sub

Private Sub FileManipulationTest()
    Dim l As Long
    Open "file-wb.txt" For Binary Access Write As #69
        'Put #69, , "string"
        'Put #69, , &HC0C0
        Put #69, , &HC0C0C0C0
        'Put #69, , 1.11113
    Close #69
    
    'Open "file-rb.txt" For Binary Access Read As #69
    Open "file-wb.txt" For Binary Access Read As #69
        Get #69, , l
        
        'MsgBox l & " and should be " & &HC0C0C0C0
        
        MsgBox "Fsize = " & LOF(69)
        Seek 69, 1
        
    Close #69
    Exit Sub

    
    Open "file-rwb.txt" For Binary Access Read Write As #69
    Close #69
    
    Open "file-out.txt" For Output As #69
    Close #69
    
    Open "file-in.txt" For Input As #69
    Close #69
    
    Open "file-rand.txt" For Random As #69
    Close #69

    Open "file-out-w.txt" For Output Access Write As #69
    Close #69
    
    Open "file-in-r.txt" For Input Access Read As #69
    Close #69
    
    Open "file-rand.txt" For Random As #69
    Close #69
End Sub

Private Sub InputBoxTest()
    MsgBox InputBox$("This is the caption of the inputbox", 123456789, "this should be the default value", , , "HelpFile", 6969), vbInformation + vbOKOnly
End Sub

Private Sub ToStringConversionTest()
    MsgBox CStr(True) & " " & CStr(1&) & " " & CStr(1#) & " " & CStr(1.1) & " " & CStr(3.1416949383838)
End Sub

Private Sub ObjectCreateTest()
    'Call CreateObject("NonExisting")
    Call CreateObject("Scripting.FileSystemObject").CreateTextFile("test.txt", True)
    MsgBox CreateObject("Scripting.FileSystemObject").GetSpecialFolder(0)
    MsgBox CreateObject("Scripting.FileSystemObject").FileExists("test.txt")
End Sub

Private Sub ObjectAppInteractionTest()
    MsgBox "App.Title = " & App.Title & vbCrLf & _
           "App.EXEName = " & App.EXEName & vbCrLf & _
           "App.Path = " & App.Path
End Sub

