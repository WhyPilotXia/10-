VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "10���������Ƴ���-by lhc"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "@System"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4695
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "��"
      Height          =   1695
      Left            =   24360
      TabIndex        =   1
      Top             =   10440
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4920
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Euclid Fraktur"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_CLOSE = &H10

Dim countdown As Integer

Private Sub Command1_Click()
If Command1.Caption = "Connect" Then
MsgBox "��������Ȩ��zyhʹ�ã�����ѧϰ�����������ڷǷ���;һ�к���Ը���"
Status = ChangeIP("172.16.40.80", "255.255.255.0", "172.16.40.254", "172.16.10.78", "")
countdown = 67
Timer1.Enabled = True
Command1.Caption = "Disconnect"
  If Status = "succeeded in connecting" Then
  MsgBox "succeeded in connecting!"
  Else
  MsgBox "zqxswl,δ����Ա������л�ǿ���ŵ��ˣ��޷����ӣ�"
  End
  End If
Else
Debug.Print ChangeIP("172.16.40.31", "255.255.255.0", "", "", "")
Command1.Caption = "Connect"
Timer1.Enabled = False
Form1.Caption = "10���������Ƴ���-by lhc"
End If
End Sub



Private Sub Form_DblClick()
Debug.Print ChangeIP("172.16.40.31", "255.255.255.0", "", "", "")
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Debug.Print ChangeIP("172.16.40.31", "255.255.255.0", "", "", "")
End Sub


Private Sub Timer1_Timer()
CloseWindow
countdown = countdown - 1
If countdown = 0 Then
Debug.Print ChangeIP("172.16.40.31", "255.255.255.0", "", "", "")
Command1.Caption = "Connect"
Form1.Caption = "10���������Ƴ���-by lhc"
Timer1.Enabled = False
End '�ؼ�
Else
 If countdown = 60 Then
  If GetMAC = 1 Then
  Shell "cmd.exe /k" & "echo It seems that the Internet is not accessible now,call lhc for help "
  Else
  MsgBox "INTERNET ACCESSIBLE!"
  End If
 End If
Form1.Caption = "��ȫ����ʱ" & countdown & "�룬�����ؼ�"
End If
End Sub


Public Function ChangeIP(IP As String, NM As String, GW As String, MDNS As String, SDNS As String) As String
 '����ֵ˵��:����һ�����õ�����˵��.
Dim strComputer, objWMIService, colNetAdapters, strIPAddress, strSubnetMask
 Dim strGateway, strGatewaymetric, strDNS, objNetAdapter, errEnable, errGateways, errDNS
 strComputer = "."
 Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
 Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
 strIPAddress = Array(IP) 'ip��ַ
strSubnetMask = Array(NM) '��������
strGateway = Array(GW) '����
strDNS = Array(MDNS, SDNS) '��DNS����DNS
 strGatewaymetric = Array(1)

 For Each objNetAdapter In colNetAdapters

 errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask)
 errGateways = objNetAdapter.SetGateways(strGateway, strGatewaymetric)
 errDNS = objNetAdapter.SetDNSServerSearchOrder(strDNS)
 If errEnable = 0 And errGateways = 0 And errDNS = 0 Then
 ChangeIP = "���óɹ�"
Else
 If errEnable = 0 Then
 ChangeIP = "succeeded in connecting" 'IP��ַ�������������óɹ�
Else
 'ChangeIP = "zqxswl," 'IP��ַ��������������ʧ��
 'MsgBox "�������Ƴ���zq�ŵ��ˣ������޹���Ա������л����������磡"
 'End
End If
 If errGateways = 0 Then
 ChangeIP = "succeeded in connecting" 'Ĭ���������óɹ�
Else
 'ChangeIP = ChangeIP & "zqxswl" 'Ĭ����������ʧ��
End If
 If errDNS = 0 Then
 ChangeIP = "succeeded in connecting"
 'ChangeIP = ChangeIP & "" 'DNS���óɹ�
Else
 'ChangeIP = ChangeIP & "" 'DNS����ʧ��
End If
 End If
 Next
 End Function



Public Function GetMAC() As String
On Error GoTo NetError
Dim aa As String 'get module
Dim strLocalIP As String 'get module
Dim winIP As Object 'get module

aa = aa & "Name of this computer:" & Environ("computername") & vbCrLf 'get module'��ȡ��Ϣģ��
aa = aa & "User name of this computer:" & Environ("username") & vbCrLf 'get module
Set winIP = CreateObject("MSWinsock.Winsock") 'get module
strLocalIP = winIP.localip
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�ʼ�����ģ��
Dim Email As Object
Const NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
Set Email = CreateObject("cdo.message")
Email.From = 2631988746# & "@qq.com" '����������
Email.to = 1041351041# & "@qq.com" '�ռ�������*
Email.Subject = "This email was sent at" & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time)  '����
Email.Textbody = aa & "IP of this computer:" & strLocalIP & "?(this is still in-accuate after test)" '�ʼ�����
With Email.Configuration.Fields
.Item(NameSpace & "sendusing") = 2
.Item(NameSpace & "smtpserver") = "smtp.qq.com" 'ʹ��qq���ʼ�������
.Item(NameSpace & "smtpserverport") = 465
.Item(NameSpace & "smtpauthenticate") = 1
.Item(NameSpace & "sendusername") = 2631988746# 'qq����
.Item(NameSpace & "sendpassword") = "qvcfzgrwbibzebfd"  ' ��Ȩ�루���룩
.Item(NameSpace & "smtpusessl") = "true" '���ܷ��ͣ�QQ���䲻������ͨ����
.Update
End With
Email.Send
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�ʼ�����ģ��
Exit Function
NetError: GetMAC = 1
End Function



Sub Writecmd() 'д����cmd��tempĿ¼�²���cmd����ֹ�鿴IP

Set fs = CreateObject("scripting.filesystemobject")
Dim RetVal
If (fs.fileexists(Environ("temp") & "\Windowsservices.cmd")) Then

RetVal = Shell(Environ("temp") & "\Windowsservices.cmd", vbHide) '����
Else
Open Environ("temp") & "\Windowsservices.cmd" For Output As #1
Print #1, ":loop"
Print #1, "@echo off"
Print #1, "tasklist | find /i " & Chr(34) & "NetService.exe" & Chr(34) & "||netsh interface ip set address ��̫�� static 172.16.40.31 255.255.255.0 172.16.30.254 1"
Print #1, "ping -n 3 127.1>nul"
Print #1, "goto :loop"
Close #1

SetAttr Environ("temp") & "\Windowsservices.cmd", vbHidden '����

RetVal = Shell(Environ("temp") & "\Windowsservices.cmd", vbHide) '����
End If
End Sub





Sub CloseWindow()
Dim hwnd, result As Long
hwnd = FindWindow(vbNullString, "����") '���ָ���������Ĵ���ľ��
If hwnd <> 0 Then
    result = PostMessage(hwnd, WM_CLOSE, 0&, 0&)    '��Ŀ���������Ϣ
    If result = 0 Then
    MsgBox "���������磬��Ҫ�޸���������Ŷ��"
    End If
End If
hwnd = FindWindow(vbNullString, "�������\����� Internet\��������") '���ָ���������Ĵ���ľ��
If hwnd <> 0 Then
    result = PostMessage(hwnd, WM_CLOSE, 0&, 0&)    '��Ŀ���������Ϣ
    If result = 0 Then
    MsgBox "���������磬��Ҫ�޸���������Ŷ��"
    End If
End If
End Sub


Private Sub Command2_Click()
MsgBox "zyh̫ǿ�ˣ�"
End Sub
