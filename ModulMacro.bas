Sub MacroEnabled()
UserForm1.Show
End Sub

Sub RegKeySave(i_RegKey As String, i_Value As String, Optional i_Type As String = "REG_SZ")
Dim myWS As Object
Set myWS = CreateObject("WScript.Shell")
myWS.RegWrite i_RegKey, i_Value, i_Type
End Sub

Private Sub UserForm_Initialize()
Dim XPath As String
'versi Office
TextBox1.Text = Application.version
XPath = Environ(20)
APath = Split(XPath, ";")
APath = Split(APath(0), "=")
TextBox2.Text = APath(1)
'Call test
End Sub

Private Sub BTEnable_Click()
Call RegKeySave("HKEY_CURRENT_USER\Software\Microsoft\Office\" & TextBox1.Text & "\Excel\Security\VBAWarnings", "1", "REG_DWORD")
MsgBox "Macro Enable"
End Sub

