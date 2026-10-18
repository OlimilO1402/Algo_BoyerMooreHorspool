Attribute VB_Name = "MApp"
Option Explicit
Public Haystack As String
Public Needle   As String
Public Start    As Long
Private m_CurForm As Form

Sub Main()
    AppInit
    Form1.Show
End Sub

Public Sub AppInit()
    Haystack = "ui ui ui ui ui the needle is in the haystack, go and find the needle"
    Needle = "needle"
    Start = 0
End Sub

Public Property Get AppInfo() As String
    AppInfo = App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
End Property

Public Property Set CurForm(ByVal aForm As Form)
    Set m_CurForm = aForm
End Property

Public Sub Debug_Print(ByVal s As String)
    m_CurForm.Text1 = m_CurForm.Text1 & s & vbCrLf
End Sub

