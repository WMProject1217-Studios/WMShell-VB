VERSION 5.00
Begin VB.Form Window 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   613.12
   ScaleMode       =   0  'User
   ScaleWidth      =   656.437
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picturebox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   4815
      TabIndex        =   1
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label talkbox 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�����˴��Կ�ʼ"
      BeginProperty Font 
         Name            =   "����ϸ��"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   1168
      TabIndex        =   0
      Top             =   4688
      Width           =   7009
   End
End
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Step

Private Sub Form_Load()
Step = 0
picturebox.Width = 412
picturebox.Height = 412
picturebox.Left = 120
picturebox.Top = 0
talkbox.Width = 480
talkbox.Height = 200
talkbox.Left = 80
talkbox.Top = 420
talkbox.ForeColor = RGB(0, 0, 0)
End Sub


Private Sub talkbox_Click()
Step = Step + 1
If Step = 1 Then
Retval = Playsound("E:\WMShell\RuinaPartLS\DW2.WAV")
Window.BackColor = RGB(255, 255, 255)
talkbox.ForeColor = RGB(0, 0, 0)
Window.Caption = "Ruina �϶����� ("
talkbox.Caption = "ˮ��ζ���ӳ���˲�֪�δ���ңԶ�⾰����"
ElseIf Step = 2 Then
picturebox.Picture = LoadPicture("E:\WMShell\RuinaPartLS\OldCity.bmp")
talkbox.Caption = vbCrLf & "�D�D�㿴����һ������˼��ĳ��С�"
ElseIf Step = 3 Then
talkbox.Caption = "��������Ľ�׽�����" & vbCrLf & "�������Ƶ���" & vbCrLf & "���ι�״����Ⱥ��"
ElseIf Step = 4 Then
talkbox.Caption = "�ڽֵ��ϳ����н���ʿ���ǣ�" & vbCrLf & "�ó�ǹ����ʬ��߾��ڿա���"
ElseIf Step = 5 Then
picturebox.Visible = False
talkbox.Caption = vbCrLf & "��ɫ�仯�˨D�D"
ElseIf Step = 6 Then
picturebox.Picture = LoadPicture("E:\WMShell\RuinaPartLS\RoomOfWitch.bmp")
Window.BackColor = RGB(0, 0, 0)
talkbox.ForeColor = RGB(255, 255, 255)
picturebox.Visible = True
talkbox.Caption = "��ĳ�������ڡ�" & vbCrLf & "ʯ������ĸ����컨�£���һ����⾰�¡�" & vbCrLf & "һ����Ů������У��������ĵ�����ϯ�ϡ�"
ElseIf Step = 7 Then
talkbox.Caption = vbCrLf & "���D�D ���� �D�D ���� �D�D��"
ElseIf Step = 8 Then
talkbox.Caption = vbCrLf & "���漴����Ĭ��ĳ�˵����֡�"
ElseIf Step = 9 Then
talkbox.Caption = vbCrLf & "���D�D ���� �D�D ���� �D�D��"
ElseIf Step = 10 Then
talkbox.Caption = vbCrLf & "���D�D�D�D" & Environ("username") & "�D�D�D�D��"
ElseIf Step = 11 Then
picturebox.Visible = False
talkbox.Caption = "�������ڣ�ˮ���ϵĹ⾰��ʧ�ˡ�"
ElseIf Step = 12 Then
Unload Window
End If
End Sub
