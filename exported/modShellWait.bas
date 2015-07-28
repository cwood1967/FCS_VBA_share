Attribute VB_Name = "modShellWait"

Private Const INFINITE = &HFFFF
Private Const SYNCHRONIZE = &H100000

Private Declare Sub WaitForSingleObject Lib "kernel32.dll" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long)

Private Declare Function OpenProcess Lib "kernel32.dll" _
   (ByVal dwDA As Long, ByVal bIH As Integer, ByVal dwPID As Long) As Long

Private Declare Sub CloseHandle Lib "kernel32.dll" (ByVal hObject As Long)

Public Sub RunCmd(CmdPath As String, _
   Optional WindowStyle As VbAppWinStyle = vbNormalFocus)

   Dim hProcess As Long

   On Error GoTo Err_RunCmd

   hProcess = OpenProcess(SYNCHRONIZE, 0, Shell(CmdPath, WindowStyle))

   If hProcess Then
       WaitForSingleObject hProcess, INFINITE
       CloseHandle hProcess
   End If

   Exit Sub

Err_RunCmd:

   err.Clear

End Sub

Public Sub testd()
    RunCmd "notepad", vbNormalFocus
    MsgBox "done"
End Sub
