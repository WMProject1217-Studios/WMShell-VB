VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "WMProject1217 Shell"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10740
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ProgramTimer 
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer StartTimeTimer 
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton RunButton 
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox CommandBox 
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Text            =   "CommandBox"
      Top             =   5280
      Width           =   9255
   End
   Begin VB.TextBox LogWindow 
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "MainForm.frx":21FD4
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function ConsoleOutput(ByVal OutputType As String, ByVal OutputText As String)
If OutputType = "log" Then
    If LogWindow.Text <> "" Then
        LogWindow.Text = LogWindow.Text & vbCrLf & "[" & Date & " " & Time & "]" & OutputText
    Else
        LogWindow.Text = "[" & Date & " " & Time & "]" & OutputText
    End If
ElseIf OutputType = "direct" Then
    If LogWindow.Text <> "" Then
        LogWindow.Text = LogWindow.Text & vbCrLf & OutputText
    Else
        LogWindow.Text = OutputText
    End If
Else
    If LogWindow.Text <> "" Then
        LogWindow.Text = LogWindow.Text & vbCrLf & "[" & Date & " " & Time & "]" & "[FAIL]Function ConsoleOutput ���� OutputType ����ȷ,�ò���ӦΪ log �� direct"
    Else
        LogWindow.Text = "[" & Date & " " & Time & "]" & "[FAIL]Function ConsoleOutput ���� Outputtype ����ȷ,�ò���ӦΪ log �� direct"
    End If
End If
LogWindow.SelStart = Len(LogWindow.Text)
End Function

Public Function ConsoleClear()
LogWindow.Text = ""
End Function

Public Function ConsoleVersion()
Retval = ConsoleOutput("direct", "WMProject1217 Shell [Insider Preview][Version 0.92.443]")
Retval = ConsoleOutput("direct", "WMProject1217 Studios")
End Function

Private Sub CommandBox_KeyPress(keyascii As Integer)
If keyascii = 13 Then
RunButton_Click
End If
End Sub

Private Sub Form_Load()
SYS_Value_StartTime = 0
StartTimeTimer.Interval = 32
ProgramTimer.Interval = 32
StartTimeTimer.Enabled = True
RunButton.Enabled = False
CommandBox.Text = ""
Retval = ConsoleClear()
Retval = ConsoleVersion()
Retval = ConsoleOutput("log", "Initializing system......")
Retval = ConsoleOutput("log", "Path is '" & App.Path & "'")
Retval = ConsoleOutput("log", "Runs on " & GetWindowsVersion())
Retval = -114514
Do While Retval < 1919
DoEvents
Retval = Retval + 1
Loop
Retval = SYS_Value_StartTime
If Retval = Int(Retval) Then
    Retval = Retval & ".00"
End If
Retval = ConsoleOutput("log", "Done at " & (SYS_Value_StartTime / 31.25) & "s")
RunButton.Enabled = True
End Sub

Private Sub Form_Resize()
On Error GoTo Error
LogWindow.Top = 0
LogWindow.Left = 0
LogWindow.Width = MainForm.ScaleWidth
LogWindow.Height = MainForm.ScaleHeight - CommandBox.Height - 100
CommandBox.Top = LogWindow.Height
CommandBox.Left = 0
CommandBox.Width = MainForm.ScaleWidth - RunButton.Width - 100
CommandBox.Height = 500
RunButton.Top = CommandBox.Top
RunButton.Left = CommandBox.Width + 30
RunButton.Width = 1340
RunButton.Height = 500
Error:
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub RunButton_Click()
SYS_Value_PingTime = 0
ProgramTimer.Enabled = True
RunButton.Enabled = False
CommandBox.Enabled = False
SYS_Temp_Command = CommandBox.Text
Retval = ConsoleOutput("direct", ">" & SYS_Temp_Command)
'ִ�б�������

If SYS_Temp_Command = "Exit" Or SYS_Temp_Command = "exit" Or SYS_Temp_Command = "EXIT" Then
    End
ElseIf Left(SYS_Temp_Command, 5) = "Echo " Or Left(SYS_Temp_Command, 5) = "echo " Or Left(SYS_Temp_Command, 5) = "ECHO " Then
    Retval = ConsoleOutput("direct", Right(SYS_Temp_Command, Len(SYS_Temp_Command) - 5))
ElseIf Left(SYS_Temp_Command, 6) = "Echol " Or Left(SYS_Temp_Command, 6) = "echol " Or Left(SYS_Temp_Command, 6) = "ECHOL " Then
    Retval = ConsoleOutput("log", Right(SYS_Temp_Command, Len(SYS_Temp_Command) - 5))
ElseIf SYS_Temp_Command = "Clean" Or SYS_Temp_Command = "clean" Or SYS_Temp_Command = "CLEAN" Or SYS_Temp_Command = "cls" Or SYS_Temp_Command = "CLS" Then
    Retval = ConsoleClear()
ElseIf Left(SYS_Temp_Command, 5) = "Exec " Or Left(SYS_Temp_Command, 5) = "exec " Or Left(SYS_Temp_Command, 5) = "EXEC " Then
    SYS_TEMP_RETURN = Execute("Normal", Right(SYS_Temp_Command, Len(SYS_Temp_Command) - 5))
    If SYS_TEMP_RETURN <> 0 Then
        Retval = ConsoleOutput("log", "[FAIL]����Ӧ�ó���ʧ��: ���� " & SYS_TEMP_RETURN)
    Else
        Retval = ConsoleOutput("log", "[INFO]������ " & Right(SYS_Temp_Command, Len(SYS_Temp_Command) - 5))
    End If
ElseIf Left(SYS_Temp_Command, 7) = "Exechc " Or Left(SYS_Temp_Command, 7) = "exechc " Or Left(SYS_Temp_Command, 7) = "EXECHC " Then
    SYS_TEMP_RETURN = Execute("Hide", Right(SYS_Temp_Command, Len(SYS_Temp_Command) - 5))
    If SYS_TEMP_RETURN <> 0 Then
        Retval = ConsoleOutput("log", "[FAIL]����Ӧ�ó���ʧ��: ���� " & SYS_TEMP_RETURN)
    Else
        Retval = ConsoleOutput("log", "[INFO]������ " & Right(SYS_Temp_Command, Len(SYS_Temp_Command) - 7))
    End If
ElseIf Left(SYS_Temp_Command, 10) = "Playsound " Or Left(SYS_Temp_Command, 10) = "playsound " Or Left(SYS_Temp_Command, 10) = "PLAYSOUND " Then
    SYS_TEMP_RETURN = Playsound(Right(SYS_Temp_Command, Len(SYS_Temp_Command) - 10))
    If SYS_TEMP_RETURN <> 0 Then
        Retval = ConsoleOutput("log", "[FAIL]Playsound Failure " & SYS_TEMP_RETURN)
    Else
        Retval = ConsoleOutput("log", "[INFO]�� " & Environ("username") & " ���� " & ShortPath(Right(SYS_Temp_Command, Len(SYS_Temp_Command) - 10)))
    End If
ElseIf Left(SYS_Temp_Command, 9) = "Countpai " Or Left(SYS_Temp_Command, 9) = "countpai " Or Left(SYS_Temp_Command, 9) = "COUNTPAI " Then
    Dim n As Double, t As Double, s As Double, m As Double
    m = Val(Right(SYS_Temp_Command, Len(SYS_Temp_Command) - 9))
    s = 2#
    For n = 1 To m
        DoEvents
        t = (2 * n) ^ 2 / ((2 * n - 1) * (2 * n + 1))
        s = s * t
    Next
    Retval = ConsoleOutput("log", "�еĽ���ֵΪ : " & s)
ElseIf SYS_Temp_Command = "DebugW" Then
    Load Window
    Window.Show
ElseIf SYS_Temp_Command = "Help" Or SYS_Temp_Command = "help" Or SYS_Temp_Command = "HELP" Then
    SYS_TEMP_RETURN = "����" & vbCrLf
    SYS_TEMP_RETURN = SYS_TEMP_RETURN & "Exit �˳����ն�" & vbCrLf
    SYS_TEMP_RETURN = SYS_TEMP_RETURN & "Echo [string] ����ı�������̨" & vbCrLf
    SYS_TEMP_RETURN = SYS_TEMP_RETURN & "Echol [string] ����־��ʽ����ı�������̨" & vbCrLf
    SYS_TEMP_RETURN = SYS_TEMP_RETURN & "Clean ��տ���̨" & vbCrLf
    SYS_TEMP_RETURN = SYS_TEMP_RETURN & "Exec [path] ����Ӧ�ó���" & vbCrLf
    SYS_TEMP_RETURN = SYS_TEMP_RETURN & "Exechc [path] �Ժ�̨������������̨Ӧ�ó���" & vbCrLf
    SYS_TEMP_RETURN = SYS_TEMP_RETURN & "Playsound [path] �ڵ�ǰ�����ϲ�����Ƶ�ļ�" & vbCrLf
    SYS_TEMP_RETURN = SYS_TEMP_RETURN & "Countpai [value] ��ָ���ļ�����������" & vbCrLf
    SYS_TEMP_RETURN = SYS_TEMP_RETURN & "DebugW �����ʵ�" & vbCrLf
    SYS_TEMP_RETURN = SYS_TEMP_RETURN & "Status �鿴������Ϣ" & vbCrLf
    SYS_TEMP_RETURN = SYS_TEMP_RETURN & "Ping ����ִ��if������Ҫ��ʱ��"
    Retval = ConsoleOutput("direct", SYS_TEMP_RETURN)
ElseIf SYS_Temp_Command = "Status" Or SYS_Temp_Command = "status" Or SYS_Temp_Command = "STATUS" Then
    Dim MUT As MEMORY_USAGE
    Dim CPUInfo, MemoryInfo, DiskInfo
    Retval = GetMemoryInfo(MUT)
    CPUInfo = "CPUռ�� : " & GetCPUUsage()
    MemoryInfo = "����RAM : " & ConvertByteNumber(MUT.AvailablePhysicalMemory) & "/" & ConvertByteNumber(MUT.PhysicalMemorySize)
    DiskInfo = "���ô��̿ռ� : " & GetDiskInfo()
    Retval = ConsoleOutput("direct", "Ӧ�ó���汾 : [Insider Preview]0.92.443" & vbCrLf & "ϵͳ�汾 : " & GetWindowsVersion() & vbCrLf & "������ʱ�� : " & (SYS_Value_StartTime / 31.25) & "s" & vbCrLf & CPUInfo & vbCrLf & MemoryInfo & vbCrLf & DiskInfo)
ElseIf SYS_Temp_Command = "Ping" Or SYS_Temp_Command = "ping" Or SYS_Temp_Command = "PING" Then
    Retval = SYS_Value_PingTime / 31.25
    Retval = ConsoleOutput("log", "ž! [" & Retval & "s]")
Else
    Retval = ConsoleOutput("log", "'" & SYS_Temp_Command & "' " & " ���ܲ�����ȷ��ָ��Ŷ?")
End If
CommandBox.Enabled = True
RunButton.Enabled = True
ProgramTimer.Enabled = False
CommandBox.SelStart = Len(CommandBox.Text)
CommandBox.SetFocus
End Sub

Private Sub StartTimeTimer_Timer()
SYS_Value_StartTime = SYS_Value_StartTime + 1
End Sub

Private Sub ProgramTimer_Timer()
SYS_Value_PingTime = SYS_Value_PingTime + 1
End Sub

