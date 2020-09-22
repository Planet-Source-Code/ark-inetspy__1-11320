VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Spy Settings"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4455
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Run at StartUp"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Spy Log File"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Inet Spy"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IEWin As cIEWindows

Private Sub Check1_Click()
  If Check1.Value = GetSetting("InetSpy", "Settings", "RunAtStartUp", 0) Then
     Command3.Enabled = False
  Else
     Command3.Enabled = True
  End If
End Sub

Private Sub Command1_Click()
   If Command1.Caption = "Activate" Then
      Set IEWin = New cIEWindows
      AddToLog vbCrLf & vbTab & vbTab & vbTab & "Spy Activated " & Now
      Command1.Caption = "Deactivate"
      Label2 = "Status: Active"
      Command2.Caption = "Hide"
   Else
      Set IEWin = Nothing
      AddToLog vbTab & vbTab & vbTab & "Spy Deactivated " & Now
      Command1.Caption = "Activate"
      Label2 = "Status: Passive"
      Command2.Caption = "Close"
   End If
End Sub

Private Sub Command2_Click()
  If Command2.Caption = "Close" Then
     Unload Me
  Else
     Me.Hide
  End If
End Sub

Private Sub Command3_Click()
  Dim objWSHShell As Object
  Set objWSHShell = CreateObject("WScript.Shell")
  Command3.Enabled = False
  Call SaveSetting("InetSpy", "Settings", "LogFile", Text1.Text)
  Call SaveSetting("InetSpy", "Settings", "RunAtStartUp", CStr(Check1.Value))
  sFile = Text1
  If Check1.Value = 1 Then
     objWSHShell.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\InetSpy", App.Path & "\" & "InetSpy.exe"
  Else
     On Error Resume Next
     objWSHShell.RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\GetSound"
     On Error GoTo 0
  End If
  Set objWSHShell = Nothing
End Sub

Private Sub Form_Load()
   App.TaskVisible = False
   Me.Hide
   Command3.Caption = "Apply"
   Command3.Enabled = False
   Label2 = "Status: Active"
   Command1.Caption = "Deactivate"
   Command2.Caption = "Hide"
   Text1 = GetSetting("InetSpy", "Settings", "LogFile", "c:\Spylog.txt")
   Check1.Value = GetSetting("InetSpy", "Settings", "RunAtStartUp", 0)
   sFile = Text1
   Set IEWin = Nothing
   Set IEWin = New cIEWindows
   AddToLog vbCrLf & vbTab & vbTab & vbTab & "Spy Activated " & Now & vbCrLf
   SetHotKey hwnd, MOD_CONTROL + MOD_SHIFT, vbKeyP
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RemoveHotKey
  Set IEWin = Nothing
End Sub

Private Sub Text1_Change()
   If Text1 = GetSetting("InetSpy", "Settings", "LogFile", "c:\Spylog.txt") Then
      Command3.Enabled = False
   Else
     Command3.Enabled = True
   End If
End Sub
