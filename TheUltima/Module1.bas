Attribute VB_Name = "Protect"
Public Type DEBUG_EVENT
dwDebugEventCode As Long
dwProcessId As Long
dwThreadId As Long
DATA(20) As Long 'enough space
End Type

Declare Function SetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function OpenEvent Lib "kernel32" Alias "OpenEventA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As Any, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const EVENT_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3)

Public Const SW_SHOWNORMAL = 1
Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Public Const INFINITE = &HFFFFFFFF
Public Const DBG_CONTINUE = &H10002
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const EXIT_PROCESS_DEBUG_EVENT = 5
Declare Function ContinueDebugEvent Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwThreadId As Long, ByVal dwContinueStatus As Long) As Long
Declare Function WaitForDebugEvent Lib "kernel32.dll" (ByRef lpDebugEvent As DEBUG_EVENT, ByVal dwMilliseconds As Long) As Long
Declare Function DebugActiveProcess Lib "kernel32" (ByVal dwProcessId As Long) As Long
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, lpCommandLine As Any, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As Any, lpProcessInformation As Any) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Sub Main()
'This example working only with compiled exe...
'when you finish your project,bulit-in this protection as the final thing !
Dim ret As Long
Select Case Command

Case ""
  Dim PIP As PROCESS_INFORMATION
  Dim SC1 As SECURITY_ATTRIBUTES
  Dim SC2 As SECURITY_ATTRIBUTES
  Dim SI As STARTUPINFO
  SC1.nLength = Len(SC1)
  SC2.nLength = Len(SC2)
  SI.cb = Len(sinfo)
  SI.dwFlags = 1
  SI.wShowWindow = 1



Dim RunAgain As String
RunAgain = AddSlash(App.Path) & App.EXEName & ".exe Zastita"
ret = CreateProcess(vbNullString, ByVal RunAgain, SC1, SC2, 0, &H20, 0&, App.Path, SI, PIP)
WaitForSingleObject PIP.hProcess, 100
Protect PIP.dwProcessId


Case "Zastita"
'Start The Program!
'yo can also crypt COMMAND LINE for the better protection!
Dim EV As Long
EV = CreateEvent(ByVal 0&, 0, 0, "GO!")
Call WaitForSingleObject(EV, INFINITE)
CloseHandle EV
Form1.Show

End Select

End Sub

Function AddSlash(ByVal StrX As String) As String
If Right(StrX, 1) <> "\" Then AddSlash = StrX & "\"
End Function


Sub Protect(ByVal ProcessId As Long)
Call DebugActiveProcess(ProcessId)
Dim EVX As Long
EVX = OpenEvent(EVENT_ALL_ACCESS, 0, "GO!")
SetEvent EVX
Dim DBGEVENT As DEBUG_EVENT
Do
WaitForDebugEvent DBGEVENT, INFINITE
Select Case DBGEVENT.dwDebugEventCode
Case EXIT_PROCESS_DEBUG_EVENT
Dim H As Long
H = OpenProcess(PROCESS_ALL_ACCESS, 0, DBGEVENT.dwProcessId)
TerminateProcess H, 0
CloseHandle H
End
End Select
Call ContinueDebugEvent(DBGEVENT.dwProcessId, DBGEVENT.dwThreadId, &H10002)
Loop

End Sub
