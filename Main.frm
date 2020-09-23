VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Main 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bandwidth Monitor"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3255
   FillColor       =   &H0080FF80&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   795
      Left            =   60
      ScaleHeight     =   735
      ScaleWidth      =   3075
      TabIndex        =   9
      Top             =   1020
      Width           =   3135
      Begin VB.Label lblType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   1035
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Average"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   0
         Width           =   795
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   0
         Width           =   795
      End
      Begin VB.Label lblAvrR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblAvrS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblBPSs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Received"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblBPSr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bps"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bps"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Timer tmrGraph 
      Interval        =   1000
      Left            =   3180
      Top             =   360
   End
   Begin VB.PictureBox Chart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      FillColor       =   &H00008000&
      ForeColor       =   &H00FFFF80&
      Height          =   915
      Left            =   60
      ScaleHeight     =   58
      ScaleMode       =   0  'User
      ScaleWidth      =   200
      TabIndex        =   8
      Top             =   60
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3720
      Top             =   600
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   4260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":015C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":02B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0414
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "by Mike15"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click to minimize to the system tray"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   3780
      Width           =   495
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   3540
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Received"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3795
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sent"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3555
      Width           =   975
   End
   Begin VB.Label lblSent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   3540
      Width           =   1335
   End
   Begin VB.Label lblRecv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   3780
      Width           =   1335
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private ChartSend(1 To 20) As Long
Private ChartRecv(1 To 20) As Long
Private LastSend As Long
Private LastRecv As Long
Private currSend As Long
Private currRecv As Long
Private gScale As Long
Private CurrPoss As Long
'constants required by Shell_NotifyIcon API call:
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const MAXLEN_IFDESCR = 256
Private Const MAXLEN_PHYSADDR = 8
Private Const MAX_INTERFACE_NAME_LEN = 256
Private nid As NOTIFYICONDATA
Private m_objIpHelper As CIpHelper

Private Sub Form_Click()
Me.WindowState = vbNormal
End Sub

Private Sub Form_Load()
CurrPoss = 1
gScale = 1
Set m_objIpHelper = New CIpHelper
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = ImageList1.ListImages(4).Picture
nid.szTip = "Bytes received: " & Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###")) & vbCrLf & " Bytes sent: " & Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###")) & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
PlotChart
End Sub
 Private Sub PlotChart()
 Dim gCount As Long
 Dim gPos As Long
 Dim rc As Long
 Dim sc As Long
 Dim avrS As Long
 Dim avrR As Long
 Dim maxRS As Long
 Dim scval As Long
 
 
 Chart.Cls
 scval = CStr(CSng(gScale) * 25)
 Me.Caption = "Net Monitor (Full Scale = " & CStr(scval) & ")"
 For gCount = 1 To 20
    gPos = CurrPoss + gCount - 2
    gPos = gPos Mod 20 + 1
    Chart.Line ((gCount * 10) - 8, 27)-((gCount * 10) - 2, 27 + (ChartSend(gPos)) / gScale), vbCyan, BF
    Chart.Line ((gCount * 10) - 8, 30)-((gCount * 10) - 2, 30 - (ChartRecv(gPos)) / gScale), &H80C0FF, BF
    If ChartSend(gPos) <> 0 Then
        sc = sc + 1
        avrS = avrS + ChartSend(gPos)
        If maxRS > ChartSend(gPos) Then maxRS = ChartSend(gPos)
    End If
    If ChartRecv(gPos) <> 0 Then
        rc = rc + 1
        avrR = avrR + ChartRecv(gPos)
        If maxRS > ChartRecv(gPos) Then maxRS = ChartRecv(gPos)
    End If
 Next gCount
 If sc <> 0 Then Me.lblAvrS = -(CLng(avrS / sc)) Else Me.lblAvrS = 0
 If rc <> 0 Then Me.lblAvrR = -(CLng(avrR / rc)) Else Me.lblAvrR = 0
 Chart.ForeColor = vbBlue
 Chart.Line (0, 28)-(200, 28)
 Chart.Line (0, 27 - 25)-(200, 27 - 25)
 Chart.Line (0, 30 + 25)-(200, 30 + 25)

 Chart.DrawStyle = 2
 Chart.Line (0, 27 - 12)-(200, 27 - 12)
 Chart.Line (0, 30 + 12)-(200, 30 + 12)
 Chart.DrawStyle = 0
 Chart.ForeColor = vbBlack
 Chart.Line (0, 29)-(200, 29)
 Chart.Line (0, 27)-(200, 27)
    gScale = -(maxRS / 25) + 1

 End Sub
Private Sub UpdateInterfaceInfo()
Dim objInterface        As CInterface
Static st_objInterface  As CInterface
Static lngBytesRecv     As Long
Static lngBytesSent     As Long
Dim blnIsRecv           As Boolean
Dim blnIsSent           As Boolean
If st_objInterface Is Nothing Then Set st_objInterface = New CInterface
Set objInterface = m_objIpHelper.Interfaces(1)
Select Case objInterface.InterfaceType
Case MIB_IF_TYPE_ETHERNET: lblType.Caption = "Ethernet"
Case MIB_IF_TYPE_FDDI: lblType.Caption = "FDDI"
Case MIB_IF_TYPE_LOOPBACK: lblType.Caption = "Loopback"
Case MIB_IF_TYPE_OTHER: lblType.Caption = "Other"
Case MIB_IF_TYPE_PPP: lblType.Caption = "PPP"
Case MIB_IF_TYPE_SLIP: lblType.Caption = "SLIP"
Case MIB_IF_TYPE_TOKENRING: lblType.Caption = "TokenRing"
End Select
lblRecv.Caption = Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###"))
lblSent.Caption = Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###"))
Set st_objInterface = objInterface
'---------------
blnIsRecv = (m_objIpHelper.BytesReceived > lngBytesRecv)
blnIsSent = (m_objIpHelper.BytesSent > lngBytesSent)
If blnIsRecv And blnIsSent Then
nid.hIcon = ImageList1.ListImages(4).Picture
ElseIf (Not blnIsRecv) And blnIsSent Then
nid.hIcon = ImageList1.ListImages(3).Picture
ElseIf blnIsRecv And (Not blnIsSent) Then
nid.hIcon = ImageList1.ListImages(2).Picture
ElseIf Not (blnIsRecv And blnIsSent) Then
nid.hIcon = ImageList1.ListImages(1).Picture
End If
lngBytesRecv = m_objIpHelper.BytesReceived
lngBytesSent = m_objIpHelper.BytesSent
nid.szTip = "Bytes received: " & Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###")) & vbCrLf & " Bytes sent: " & Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###")) & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.WindowState = vbNormal

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbBlack
Dim Result As Long
Dim msg As Long
If Me.ScaleMode = vbPixels Then
msg = X
Else
msg = X / Screen.TwipsPerPixelX
End If
Select Case msg
Case WM_LBUTTONDBLCLK    '515 restore form window
If Me.WindowState = vbNormal Then
    Me.WindowState = vbMinimized
    Me.Visible = False
Else
    Me.WindowState = vbNormal
    Me.Visible = True
    Me.SetFocus
End If
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbBlack
End Sub
Private Sub lblRecv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbBlack
End Sub
Private Sub Label5_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbBlue
End Sub

Private Sub Timer1_Timer()
Call UpdateInterfaceInfo
End Sub

Private Sub tmrGraph_Timer()
Dim bs, br As Long
bs = m_objIpHelper.BytesSent
br = m_objIpHelper.BytesReceived
If LastRecv = 0 Then LastRecv = br
If LastSend = 0 Then LastSend = bs
currRecv = LastRecv - br
currSend = LastSend - bs
LastRecv = br
LastSend = bs
ChartSend(CurrPoss) = currSend
ChartRecv(CurrPoss) = currRecv
lblBPSr = -(currRecv)
lblBPSs = -(currSend)
If CurrPoss = 20 Then
    CurrPoss = 1
Else
    CurrPoss = CurrPoss + 1
End If
    PlotChart
End Sub
