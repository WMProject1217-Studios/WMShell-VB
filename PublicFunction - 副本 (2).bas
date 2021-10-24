Attribute VB_Name = "PublicFunction"
Option Explicit

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Public Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Public Type PROCESS_QUOTA_LIMITS
    PagedPoolLimit As Long
    NonPagedPoolLimit As Long
    MinimumWorkingSetSize As Long
    MaximumWorkingSetSize As Long
    PagefileLimit As Long
    TimeLimit As LARGE_INTEGER
    Unknown As Long
End Type

Public Type SYSTEM_BASIC_INFORMATION
    Reserved As Long
    TimerResolution As Long
    PageSize As Long
    NumberOfPhysicalPages As Long
    LowestPhysicalPage As Long
    HighestPhysicalPage As Long
    AllocationGranularity As Long
    MinimumUserModeAddress As Long
    MaximumUserModeAddress As Long
    ActiveProcessorsAffinityMask As Long
    NumberOfProcessors As Byte
End Type

Public Type VM_COUNTERS
    PeakVirtualSize As Long
    VirtualSize As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Public Type MEMORY_USAGE
    LoadedMemory As Long
    PhysicalMemorySize As Currency
    AvailablePhysicalMemory As Currency
    PagedPoolSize As Currency
    NonPagedPoolSize As Currency
    PagefileMemorySize As Currency
    AvailablePagefileMemory As Currency
    VirtualMemorySize As Currency
    AvailableVirtualMemory As Currency
End Type

Public Type CRITICAL_SECTION
    DebugInfo As Long
    LockCount As Long
    RecursionCount As Long
    OwningThread As Long
    Reserved As Long
End Type

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVallpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function NtQueryInformationProcess Lib "ntdll.dll" (ByVal ProcessHandle As Long, ByVal ProcessInformationClass As Long, ByVal ProcessInformation As Long, ByVal ProcessInformationLength As Long, ByRef ReturnLength As Long) As Long
Public Declare Function NtQuerySystemInformation Lib "ntdll.dll" (ByVal SystemInformationClass As Long, ByVal SystemInformation As Long, ByVal SystemInformationLength As Long, ByRef ReturnLength As Long) As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateThreadL Lib "kernel32" Alias "CreateThread" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Sub ExitThread Lib "kernel32" (Optional ByVal dwExitCode As Long = 0)
Public Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Const CREATE_SUSPENDED = &H4
Public Retval
Public SYS_Value_StartTime, SYS_Value_PingTime
Public SYS_Temp_Command, SYS_Temp_Function, SYS_Temp_LoadScript
Public SYS_TEMP_RETURN

Public Function ShortPath(ByVal LongPath As String) As String
Dim tmpStr As String * 255, intLnth As Integer
intLnth = GetShortPathName(LongPath, tmpStr, 255)
ShortPath = Left$(tmpStr, intLnth)
End Function

Public Function ConvertByteNumber(ByVal ByteNumber As Currency)
If ByteNumber < 1024 Then
    ConvertByteNumber = ByteNumber & " B"
    Exit Function
Else
    ByteNumber = ByteNumber / 1024
    If ByteNumber < 1024 Then
        ConvertByteNumber = ByteNumber & " KB"
        Exit Function
    Else
        ByteNumber = ByteNumber / 1024
        If ByteNumber < 1024 Then
            ConvertByteNumber = ByteNumber & " MB"
            Exit Function
        Else
            ByteNumber = ByteNumber / 1024
            If ByteNumber < 1024 Then
                ConvertByteNumber = ByteNumber & " GB"
                Exit Function
            Else
                ByteNumber = ByteNumber / 1024
                If ByteNumber < 1024 Then
                    ConvertByteNumber = ByteNumber & " TB"
                    Exit Function
                Else
                    ByteNumber = ByteNumber / 1024
                    If ByteNumber < 1024 Then
                        ConvertByteNumber = ByteNumber & " PB"
                        Exit Function
                    Else
                        ByteNumber = ByteNumber / 1024
                        ConvertByteNumber = ByteNumber & " EB"
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
End If
End Function

Public Function Playsound(ByVal AudioPath As String)
On Error GoTo Error
Retval = mciSendString("CLOSE BackgroundMusic", "", 0, 0)
Retval = mciSendString("OPEN " & ShortPath(AudioPath) & " ALIAS BackgroundMusic", "", 0, 0)
Retval = mciSendString("PLAY BackgroundMusic FROM 0", "", 0, 0)
Playsound = 0
Exit Function
Error:
Playsound = 1145
End Function

Public Function Execute(ByVal LauncherType As String, ByVal FullPath As String)
On Error GoTo Error
If LauncherType = "" Then
    Retval = Shell(FullPath, vbNormalFocus)
ElseIf LauncherType = "Normal" Then
    Retval = Shell(FullPath, vbNormalFocus)
ElseIf LauncherType = "Hide" Then
    Retval = Shell(FullPath, vbHide)
Else
    GoTo Error
End If
Execute = 0
Exit Function
Error:
Execute = 1145
End Function

Public Function GetCPUUsage()
Dim objProc As Object
Set objProc = GetObject("winmgmts:\\.\root\cimv2:win32_processor='cpu0'")
GetCPUUsage = Int(objProc.LoadPercentage) & "%"
End Function

Public Function GetMemoryInfo(ByRef MemoryUsage As MEMORY_USAGE)
Dim BasicInfo As SYSTEM_BASIC_INFORMATION
Dim QuotaLimits As PROCESS_QUOTA_LIMITS
Dim Vm As VM_COUNTERS
Dim PerInfo(77) As Long
Dim Status As Long
Dim ReturnLength As Long
Dim ProcessAPagefile As Currency
Dim SystemAPagefile As Currency
Status = NtQuerySystemInformation(0, ByVal VarPtr(BasicInfo), LenB(BasicInfo), 0)
If Status Then Exit Function
Status = NtQuerySystemInformation(2, ByVal VarPtr(PerInfo(0)), 312, 0)
If Status Then Exit Function
Status = NtQueryInformationProcess(-1, 1, ByVal VarPtr(QuotaLimits), Len(QuotaLimits), ReturnLength)
Status = NtQueryInformationProcess(-1, 3, ByVal VarPtr(Vm), LenB(Vm), 0)
If PerInfo(11) < 100 Then
    MemoryUsage.LoadedMemory = 100
Else
    MemoryUsage.LoadedMemory = Fix((BasicInfo.NumberOfPhysicalPages - PerInfo(11)) * 100 / BasicInfo.NumberOfPhysicalPages)
End If
MemoryUsage.PhysicalMemorySize = CCur(BasicInfo.NumberOfPhysicalPages) * BasicInfo.PageSize
MemoryUsage.AvailablePhysicalMemory = CCur(PerInfo(11)) * BasicInfo.PageSize
MemoryUsage.PagedPoolSize = CCur(PerInfo(28)) * BasicInfo.PageSize
MemoryUsage.NonPagedPoolSize = CCur(PerInfo(29)) * BasicInfo.PageSize
If PerInfo(13) < QuotaLimits.PagefileLimit Then
    MemoryUsage.PagefileMemorySize = CCur(QuotaLimits.PagefileLimit) * BasicInfo.PageSize
Else
    MemoryUsage.PagefileMemorySize = CCur(PerInfo(13)) * BasicInfo.PageSize
End If
MemoryUsage.VirtualMemorySize = BasicInfo.MaximumUserModeAddress - BasicInfo.MinimumUserModeAddress + 1
MemoryUsage.AvailableVirtualMemory = MemoryUsage.VirtualMemorySize - Vm.VirtualSize
SystemAPagefile = PerInfo(13) - PerInfo(12)
ProcessAPagefile = QuotaLimits.PagefileLimit - Vm.PagefileUsage
Status = Fix(MemoryUsage.AvailablePhysicalMemory / MemoryUsage.PhysicalMemorySize * 100)
If SystemAPagefile > ProcessAPagefile Then
    MemoryUsage.AvailablePagefileMemory = CCur(SystemAPagefile) * BasicInfo.PageSize
Else
    MemoryUsage.AvailablePagefileMemory = CCur(ProcessAPagefile) * BasicInfo.PageSize
End If
End Function

Public Function GetDiskInfo()
Dim UserBytes As LARGE_INTEGER, TotalBytes As LARGE_INTEGER, FreeBytes As LARGE_INTEGER
Dim mUserBytes As Double, mTotalBytes As Double
Dim ctlNew As Control
Dim i As Integer, DriveName As String
mTotalBytes = 0
mUserBytes = 0
Set ctlNew = MainForm.Controls.Add("VB.drivelistbox", "cmdNew", MainForm)
With ctlNew
    For i = 0 To .ListCount - 1
        DriveName = Left(.List(i), InStr(.List(i), ":"))
        If GetDriveType(DriveName) = 3 Then
            Retval = GetDiskFreeSpaceEx(DriveName & "\", UserBytes, TotalBytes, FreeBytes)
            mTotalBytes = mTotalBytes + (TotalBytes.HighPart * (16 ^ 8) + TotalBytes.LowPart + IIf(TotalBytes.LowPart < 0, (16 ^ 8), 0))
            mUserBytes = mUserBytes + (UserBytes.HighPart * (16 ^ 8) + UserBytes.LowPart + IIf(UserBytes.LowPart < 0, (16 ^ 8), 0))
        End If
    Next
End With
MainForm.Controls.Remove ctlNew
GetDiskInfo = ConvertByteNumber(mUserBytes) & "/" & ConvertByteNumber(mTotalBytes)
End Function

Public Function GetWindowsVersion()
Dim OS As String
Dim ver As OSVERSIONINFO, retLng As Long
ver.dwOSVersionInfoSize = Len(ver)
GetVersionEx ver
With ver
    Select Case .dwPlatformId
    Case 1
        Select Case .dwMinorVersion
        Case 0
            Select Case .szCSDVersion
            Case "C"
                OS = "Windows 95 OSR2"
            Case "B"
                OS = "Windows 95 OSR2"
            Case Else
                OS = "Windows 95"
            End Select
        Case 10
            Select Case .szCSDVersion
            Case "A"
                OS = "Windows 98 SE"
            Case Else
                OS = "Windows 98"
            End Select
        Case 90
            OS = "Windows Millennium Edition"
        End Select
    Case 2
        Select Case .dwMajorVersion
        Case 3
            OS = "Windows NT 3.51"
        Case 4
            OS = "Windows NT 4.0"
        Case 5
            Select Case .dwMinorVersion
            Case 0
                Select Case .wSuiteMask
                Case &H80
                    OS = "Windows 2000 Data Center"
                Case &H2
                    OS = "Windows 2000 Advanced"
                Case Else
                    OS = "Windows 2000"
                End Select
            Case 1
                Select Case .wSuiteMask
                Case &H0
                        OS = "Windows XP Professional"
                Case &H200
                        OS = "Windows XP Home"
                Case Else
                        OS = "Windows XP"
                End Select
            Case 2
                Select Case .wSuiteMask
                Case &H2
                    OS = "Windows Server 2003 Enterprise"
                Case &H80
                    OS = "Windows Server 2003 Data Center"
                Case &H400
                    OS = "Windows Server 2003 Web Edition"
                Case &H0
                    OS = "Windows Server 2003 Standard"
                Case &H112
                    OS = "Windows Server 2003 R2 Enterprise"
                Case Else
                    OS = "Windows Server 2003"
                End Select
            End Select
            If .wServicePackMajor > 0 Then
                OS = OS & " Service Pack " & .wServicePackMajor & IIf(.wServicePackMinor > 0, "." & .wServicePackMinor, vbNullString)
            End If
        Case 6
            Select Case .wProductType
            Case &H6
                OS = "Business Edition"
            Case &H10
                OS = "Business Edition (N)"
            Case &H12
                OS = "Cluster Server Edition"
            Case &H8
                OS = "Server Datacenter Edition (Full Installation)"
            Case &HC
                OS = "Server Datacenter Edition (Core Installation)"
            Case &H4
                OS = "Enterprise Edition"
            Case &H1B
                OS = "Enterprise Edition (N)"
            Case &HA
                OS = "Server Enterprise Edition (Full Installation)"
            Case &HE
                OS = "Server Enterprise Edition (Core Installation)"
            Case &HF
                OS = "Server Enterprise Edition for Itanium-based Systems"
            Case &H2
                OS = "Home Basic Edition"
            Case &H5
                OS = "Home Basic Edition (N)"
            Case &H3
                OS = "Home Premium Edition"
            Case &H1A
                OS = "Home Premium Edition (N)"
            Case &H13
                OS = "Home Server Edition"
            Case &H18
                OS = "Server for Small Business Edition"
            Case &H9
                OS = "Small Business Server"
            Case &H19
                OS = "Small Business Server Premium Edition"
            Case &H7
                OS = "Server Standard Edition (Full Installation)"
            Case &HD
                OS = "Server Standard Edition (Core Installation)"
            Case &H8
                OS = "Starter Edition"
            Case &H17
                OS = "Storage Server Enterprise Edition"
            Case &H14
                OS = "Storage Server Express Edition"
            Case &H15
                OS = "Storage Server Standard Edition"
            Case &H16
                OS = "Storage Server Workgroup Edition"
            Case &H1
                OS = "Ultimate Edition"
            Case &H1C
                OS = "Ultimate Edition (N)"
            Case &H0
                OS = "An unknown product"
            Case &HABCDABCD
                OS = "Not activated product"
            Case &H11
                OS = "Web Server Edition"
            End Select
            Select Case .dwMinorVersion
            Case 0
                Select Case .wProductType
                Case 3
                    OS = "Windows Server 2008 " & OS
                Case Else
                    OS = "Windows Vista " & OS
                End Select
            Case 1
                Select Case .wProductType
                Case 3
                    OS = "Windows Server 2008 R2 " & OS
                Case Else
                    OS = "Windows 7 " & OS
                End Select
            Case 2
                Select Case .wProductType
                Case 3
                    OS = "Windows Server 2012 " & OS
                Case Else
                    OS = "Windows 8 " & OS
                End Select
            Case 3
                Select Case .wProductType
                Case 3
                    OS = "Windows Server 2012 R2 " & OS
                Case Else
                    OS = "Windows 8.1 " & OS
                End Select
            Case 4
                Select Case .wProductType
                Case 3
                    OS = "Windows Server 2016 Beta " & OS
                Case Else
                    OS = "Windows 9 " & OS
                End Select
            End Select
            If .wServicePackMajor > 0 Then
                OS = OS & " Service Pack " & .wServicePackMajor & IIf(.wServicePackMinor > 0, "." & .wServicePackMinor, vbNullString)
            End If
        Case 10
            Select Case .wProductType
            Case 3
                OS = "Windows Server 2016/2019 " & OS
            Case Else
                OS = "Windows 10 " & OS
            End Select
        Case 11
            Select Case .wProductType
            Case 3
                OS = "Windows Server 2022 " & OS
            Case Else
                OS = "Windows 11 " & OS
            End Select
        Case Else
            OS = "Unknown System Version!"
        End Select
    End Select
    OS = OS & " [Version: " & .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber & "]"
End With
GetWindowsVersion = OS
End Function

Public Function LoadScript(ByVal ScriptPath As String)
Open ScriptPath For Input As #1
Do While Not EOF(1)
    Line Input #1, SYS_Temp_LoadScript
    '
Loop
Close #1
End Function
