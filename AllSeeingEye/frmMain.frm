VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " All Seeing Eye"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSettings 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   420
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   1980
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2730
         ScaleHeight     =   225
         ScaleWidth      =   525
         TabIndex        =   29
         Top             =   1470
         Width           =   555
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "Play sound when new E-Mail arrives?"
         Height          =   195
         Left            =   1080
         TabIndex        =   27
         Top             =   1050
         Width           =   3705
      End
      Begin VB.CheckBox chkNotify 
         Caption         =   "Show Notification Window when minimized?"
         Height          =   195
         Left            =   1080
         TabIndex        =   26
         Top             =   810
         Width           =   3705
      End
      Begin VB.TextBox txtSeconds 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3180
         TabIndex        =   25
         Text            =   "30"
         Top             =   405
         Width           =   615
      End
      Begin VB.CommandButton cmdSave2 
         BackColor       =   &H00F4EADB&
         Caption         =   "OK"
         Height          =   375
         Left            =   1650
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1980
         Width           =   915
      End
      Begin VB.CommandButton cmdCancel2 
         BackColor       =   &H00F4EADB&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1980
         Width           =   915
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   3570
         Top             =   1890
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "<- Click to change"
         Height          =   195
         Left            =   3420
         TabIndex        =   30
         Top             =   1500
         Width           =   1305
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Background Color:"
         Height          =   195
         Left            =   1200
         TabIndex        =   28
         Top             =   1500
         Width           =   1425
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check E-Mail Accounts every                seconds"
         Height          =   195
         Left            =   1065
         TabIndex        =   24
         Top             =   450
         Width           =   3375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   75
         TabIndex        =   19
         Top             =   30
         Width           =   810
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   1080
         Top             =   1410
         Width           =   3735
      End
   End
   Begin VB.Frame fraEmail 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   420
      Left            =   120
      TabIndex        =   11
      Top             =   1290
      Visible         =   0   'False
      Width           =   1980
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00F4EADB&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1890
         Width           =   915
      End
      Begin VB.TextBox txtTitle 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1650
         TabIndex        =   0
         Top             =   435
         Width           =   3375
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   1650
         TabIndex        =   1
         Top             =   770
         Width           =   3375
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1650
         TabIndex        =   2
         Top             =   1105
         Width           =   3375
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00F4EADB&
         Caption         =   "OK"
         Height          =   375
         Left            =   1650
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1890
         Width           =   915
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1650
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label lblHeading 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add Account:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   75
         TabIndex        =   16
         Top             =   30
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POP3 Server:"
         Height          =   195
         Left            =   555
         TabIndex        =   15
         Top             =   815
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POP3 Username:"
         Height          =   195
         Left            =   315
         TabIndex        =   14
         Top             =   1150
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POP3 Password:"
         Height          =   195
         Left            =   345
         TabIndex        =   13
         Top             =   1485
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Title:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   375
         TabIndex        =   12
         Top             =   480
         Width           =   1155
      End
   End
   Begin VB.Timer tmrCheckMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2190
      Top             =   1290
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   2190
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraEmailOptions 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   43
      TabIndex        =   6
      Top             =   1860
      Width           =   5405
      Begin VB.CommandButton cmdOptions 
         BackColor       =   &H00E6C4D7&
         Caption         =   "Options"
         Height          =   375
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Width           =   1155
      End
      Begin VB.CommandButton cmdCheckMail 
         BackColor       =   &H00F4EADB&
         Caption         =   "Check"
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton cmdDeleteAccount 
         BackColor       =   &H00E6C4D7&
         Caption         =   "Delete Account"
         Height          =   375
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00E6C4D7&
         Caption         =   "Add Account"
         Height          =   375
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1155
      End
   End
   Begin MSComctlLib.ListView lvwEmail 
      Height          =   1485
      Left            =   52
      TabIndex        =   17
      Top             =   300
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   2619
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblEdit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double-Click to Edit"
      Height          =   195
      Left            =   4072
      TabIndex        =   20
      Top             =   45
      Width           =   1380
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail Accounts:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   73
      TabIndex        =   10
      Top             =   30
      Width           =   1620
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuFav 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim New1 As FrmMSNPopUp

Public Sub DoCheckMail(Index As Integer)
  On Error GoTo Err
  lvwEmail.ListItems(Index + 1).SubItems(1) = "Checking..."
  lvwEmail.ListItems(Index + 1).SubItems(2) = " "
  
  DoEvents
  Pause 0.5
  
  If m_Server(Index) = "" Or m_User(Index) = "" Or m_Password(Index) = "" Then Exit Sub
  Call CheckMail(Index, m_Server(Index), m_User(Index), m_Password(Index))

Err:
  tmrCheckMail.Enabled = False
  tmrCheckMail.Enabled = True
End Sub

Sub CheckMail(Index As Integer, vServer As String, Username As String, Password As String)
  DoEvents
  m_State(Index) = POP3_Connect
  
  If m_Ready(Index) = False Then Exit Sub
  m_Ready(Index) = False
  
  On Error Resume Next
  Unload Winsock1(Index)
  Load Winsock1(Index)
    
  Winsock1(Index).Close
  Winsock1(Index).Tag = Username & "|" & Password
  Winsock1(Index).LocalPort = 0
  Winsock1(Index).Connect vServer, 110
End Sub

Private Sub cmdCancel_Click()
  fraEmail.Visible = False
  txtTitle = ""
  txtUsername = ""
  txtServer = ""
  txtPassword = ""
  Adding = False
End Sub

Private Sub cmdCancel2_Click()
  fraSettings.Visible = False
  Call LoadSettings
End Sub

Sub LoadSettings()
  If Val(GetSetting("AllSeeingEye", "Settings", "Seconds")) = 0 Then
    Call SaveSetting("AllSeeingEye", "Settings", "Seconds", "30")
  End If
  txtSeconds = Val(GetSetting("AllSeeingEye", "Settings", "Seconds"))
  tmrCheckMail.Interval = Val(txtSeconds) * 1000
  
  If GetSetting("AllSeeingEye", "Settings", "PlaySound") = "" Then
    Call SaveSetting("AllSeeingEye", "Settings", "PlaySound", "1")
  End If
  chkSound.Value = Val(GetSetting("AllSeeingEye", "Settings", "PlaySound"))
  
  If GetSetting("AllSeeingEye", "Settings", "Notify") = "" Then
    Call SaveSetting("AllSeeingEye", "Settings", "Notify", "1")
  End If
  chkNotify.Value = Val(GetSetting("AllSeeingEye", "Settings", "Notify"))
  
  Picture1.BackColor = BackColor
End Sub


Private Sub cmdCheckMail_Click()
  tmrCheckMail_Timer
End Sub

Private Sub cmdDeleteAccount_Click()
  DeleteEmailAccount (lvwEmail.SelectedItem.Tag)
  Call LoadEmailAccounts
  Call tmrCheckMail_Timer
End Sub

Private Sub cmdOptions_Click()
  tmrCheckMail.Enabled = False
  Adding = True
  Call LoadBackground
  fraSettings.Visible = True
  DoEvents
End Sub

Private Sub cmdSave_Click()
  Dim RS As Recordset
  If Adding Then
    Set RS = DB.OpenRecordset("EmailAccounts")
    RS.AddNew
  Else
    Set RS = DB.OpenRecordset("Select * From EmailAccounts Where ID=" & CurrAccount)
    RS.Edit
  End If
  
  RS!Name = txtTitle
  RS!User = txtUsername
  RS!Password = Encrypt(txtPassword)
  RS!Server = txtServer
  RS.Update
  
  Call LoadEmailAccounts
  fraEmail.Visible = False
  Call cmdCheckMail_Click
End Sub

Private Sub cmdAdd_Click()
  txtTitle = ""
  txtUsername = ""
  txtServer = ""
  txtPassword = ""
  lblHeading = "Add Account:"
  tmrCheckMail.Enabled = False
  Adding = True
  fraEmail.Visible = True
  DoEvents
  txtTitle.SetFocus
End Sub

Private Sub cmdSave2_Click()
  Call SaveSetting("AllSeeingEye", "Settings", "Seconds", txtSeconds)
  tmrCheckMail.Interval = Val(txtSeconds) * 1000
  
  Call SaveSetting("AllSeeingEye", "Settings", "PlaySound", Val(chkSound))
  Call SaveSetting("AllSeeingEye", "Settings", "Notify", Val(chkNotify))
  
  Call SaveSetting("AllSeeingEye", "Settings", "BGColor", Picture1.BackColor)
  Call LoadBackground
  
  fraSettings.Visible = False
  Call cmdCheckMail_Click
End Sub

Private Sub Form_Activate()
  If Not FirstRun Then
    tmrCheckMail_Timer
    tmrCheckMail.Enabled = True
    FirstRun = True
  End If
End Sub

Private Sub Form_Load()
  If App.PrevInstance Then End
  Call InitializeSettings
End Sub

Sub InitializeSettings()
  Set DB = OpenDatabase(App.Path & "\ConsoleData.mdb")
  Call LoadSettings
  Call LoadBackground
  Call LoadEmailAccounts
  
  lvwEmail.ColumnHeaders(1).Width = 1621.5
  lvwEmail.ColumnHeaders(2).Width = 2342.5
  lvwEmail.ColumnHeaders(3).Width = 1081
  fraEmail.Move 0, 0, 5495, 2425
  fraSettings.Move 0, 0, 5495, 2425
  
  Call CreateSystemTrayIcon(Me, "All Seeing Eye")
  
  Dim tempString As String
  Dim Spot As Integer
  tempString = GetSetting("AllSeeingEye", "Settings", "Position")
  Spot = InStr(1, tempString, ",")
  If tempString <> "" Then
    Left = Val(Left$(tempString, Spot - 1))
    Top = Val(Mid$(tempString, Spot + 1))
  Else
    Top = Screen.Height - (Height + 900)
    Left = Screen.Width - Width
  End If
End Sub

Sub LoadBackground()
  If GetSetting("AllSeeingEye", "Settings", "BGColor") <> "" Then
    BackColor = Val(GetSetting("AllSeeingEye", "Settings", "BGColor"))
  End If

  Dim X As Control
  For Each X In Me
    If TypeOf X Is Frame Or TypeOf X Is CheckBox Then X.BackColor = BackColor
  Next
End Sub

Sub LoadEmailAccounts()
  lvwEmail.ListItems.Clear
  Dim lstItem As ListItem
  Dim X As Integer
  For X = Winsock1.UBound To 0 Step -1
    If X <> 0 Then Unload Winsock1(X)
  Next
  
  ReDim m_Server(0)
  ReDim m_User(0)
  ReDim m_Password(0)
  ReDim m_State(0)
  ReDim m_Title(0)
  ReDim m_Ready(0)
  m_Ready(0) = True
  
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select * From EmailAccounts Order by Name Asc")
  If RS.RecordCount = 0 Then
    tmrCheckMail.Enabled = False
    cmdDeleteAccount.Enabled = False
    cmdCheckMail.Enabled = False
    lblEdit = "Add Account To Start"
    FirstRun = True
    Exit Sub
  Else
    cmdDeleteAccount.Enabled = True
    cmdCheckMail.Enabled = True
    lblEdit = "Double-Click to Edit"
  End If
  X = 0
  
  RS.MoveFirst
  Do While Not RS.EOF
    If X <> 0 Then
      ReDim Preserve m_Server(X)
      ReDim Preserve m_User(X)
      ReDim Preserve m_Password(X)
      ReDim Preserve m_Ready(X)
      ReDim Preserve m_State(X)
      ReDim Preserve m_Title(X)
      m_Ready(X) = True
    End If
    
    m_Server(X) = RS!Server
    m_User(X) = RS!User
    m_Password(X) = Decrypt(RS!Password)
    m_Title(X) = RS!Name
    
    Set lstItem = lvwEmail.ListItems.Add(, "Account:" & X, RS!Name)
    lstItem.SubItems(1) = " "
    lstItem.SubItems(2) = " "
    lstItem.Tag = RS!ID
    
    X = X + 1
    RS.MoveNext
  Loop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If (X / Screen.TwipsPerPixelX) = STI_LBUTTONUP Then
    WindowState = vbNormal
    Visible = True
    Me.SetFocus
    inTray = False
  End If
End Sub

Private Sub Form_Resize()
  If WindowState = vbMinimized Then
    Me.Hide
    inTray = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call SaveSetting("AllSeeingEye", "Settings", "Position", Left & ", " & Top)
  Call DeleteSystemTrayIcon(Me)
End Sub

Private Sub lvwEmail_DblClick()
  Dim I As Integer
  I = Val(Mid$(lvwEmail.SelectedItem.Key, 9))
  CurrAccount = Val(lvwEmail.SelectedItem.Tag)
  
  txtTitle = m_Title(I)
  txtUsername = m_User(I)
  txtPassword = m_Password(I)
  txtServer = m_Server(I)
  
  fraEmail.Visible = True
  lblHeading = "Edit Account:"
  DoEvents
  txtTitle.SetFocus
End Sub

Private Sub Picture1_Click()
  CD1.CancelError = True
  On Error GoTo Err:
  
  CD1.ShowColor
  Picture1.BackColor = CD1.Color
  fraSettings.BackColor = CD1.Color
  chkSound.BackColor = CD1.Color
  chkNotify.BackColor = CD1.Color
  
Err:
End Sub

Private Sub tmrCheckMail_Timer()
  cmdCheckMail.Enabled = False
  Dim X As Integer
  For X = 0 To UBound(m_User)
    DoCheckMail (X)
    Do Until m_Ready(X)
      DoEvents
    Loop
    Pause 0.5
  Next
  cmdCheckMail.Enabled = True
  
  Dim Sum As Integer
  Dim Size As Double
  For X = 1 To lvwEmail.ListItems.Count
    Sum = Sum + Val(lvwEmail.ListItems(X).SubItems(1))
    Size = Size + Val(lvwEmail.ListItems(X).SubItems(2))
  Next

  Call SendNotify(Sum, Size)

End Sub


Sub SendNotify(Sum As Integer, Size As Double)
  
  Dim Text As String
  If Sum = 1 Then
    Text = "1 E-Mail, " & Size & " KB"
  Else
    Text = Sum & " E-Mails, " & Size & " KB"
  End If
  Call ModifySystemTrayIcon(Me, "All Seeing Eye" & vbCrLf & Text)
  
  If Sum > LastSum Then
    If chkNotify.Value = 1 And inTray Then Call ShowMessagesWindow(Text)
    If chkSound.Value = 1 Then PlaySound (App.Path & "\newalert.wav")
  End If
  LastSum = Sum
  
End Sub

Sub ShowMessagesWindow(Text As String)
  Set New1 = New FrmMSNPopUp
  New1.SetNumber 450
  New1.LblText.Caption = "All Seeing Eye"
  New1.LblMessage.Caption = Replace(Text, ", ", vbCrLf & vbCrLf)
  'New1.LblOptions.Caption = TxtOptions.Text
  New1.Visible = True
End Sub

Private Sub txtPassword_GotFocus()
  txtPassword.SelStart = 0
  txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    cmdSave.SetFocus
    KeyAscii = 0
  End If
End Sub


Private Sub txtSeconds_LostFocus()
  txtSeconds = Val(txtSeconds)
End Sub

Private Sub txtServer_GotFocus()
  txtServer.SelStart = 0
  txtServer.SelLength = Len(txtServer)
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtUsername.SetFocus
    KeyAscii = 0
  End If
End Sub

Private Sub txtTitle_GotFocus()
  txtTitle.SelStart = 0
  txtTitle.SelLength = Len(txtTitle)
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtServer.SetFocus
    KeyAscii = 0
  End If
End Sub

Private Sub txtUsername_GotFocus()
  txtUsername.SelStart = 0
  txtUsername.SelLength = Len(txtUsername)
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtPassword.SetFocus
    KeyAscii = 0
  End If
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  
  Dim strData               As String
  Dim Spot                  As Integer
  Dim Username              As String
  Dim Password              As String
  Dim lIndex                As Integer
  Static intMessages        As Integer
  Static intCurrentMessage  As Integer
  Static strBuffer          As String
  Static TotalSize          As Long
  Static TotalSize2         As Long
  Dim EmailNum              As Long
  
  On Error Resume Next
   
  Spot = InStr(1, Winsock1(Index).Tag, "|")
  Username = Left(Winsock1(Index).Tag, Spot - 1)
  Password = Mid(Winsock1(Index).Tag, Spot + 1)
  
  Winsock1(Index).GetData strData
  
  If Left$(strData, 1) = "+" Or m_State(Index) = POP3_TOP Then
    Select Case m_State(Index)
      Case POP3_Connect
        intMessages = 0
        intCurrentMessage = 0
        m_State(Index) = POP3_USER
        Winsock1(Index).SendData "USER " & Username & vbCrLf
        DoEvents
      Case POP3_USER
        m_State(Index) = POP3_PASS
        Winsock1(Index).SendData "PASS " & Password & vbCrLf
        DoEvents
      Case POP3_PASS
        m_State(Index) = POP3_STAT
        Winsock1(Index).SendData "STAT" & vbCrLf
        DoEvents
      Case POP3_STAT
        intMessages = Get_After_Seperator(strData, 1, " ")
        TotalSize = Get_After_Seperator(strData, 2, " ")
        
        If intMessages = 1 Then
          lvwEmail.ListItems(Index + 1).SubItems(1) = "1 E-Mail"
        Else
          lvwEmail.ListItems(Index + 1).SubItems(1) = intMessages & " E-Mails"
        End If
        lvwEmail.ListItems(Index + 1).SubItems(2) = Format(TotalSize / 1000, "0.00") & " KB"
                
        DoEvents
        If intMessages = 0 Then
          Winsock1(Index).SendData "QUIT" & vbCrLf
          DoEvents
          m_Ready(Index) = True
          Exit Sub
        End If
        
        Winsock1(Index).SendData "QUIT" & vbCrLf
        DoEvents
        m_State(Index) = POP3_QUIT
        m_Ready(Index) = True
      Case POP3_QUIT
        Winsock1(Index).Close
        Call DisconnectMe(Index)
        m_Ready(Index) = True
    End Select
  Else
    Winsock1(Index).Close
    lvwEmail.ListItems(Index + 1).SubItems(1) = "Error"
    m_Ready(Index) = True
  End If
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
  If Number = 10053 Then
    lvwEmail.ListItems("Account:" & Index).SubItems(1) = "Error"
    Winsock1(Index).Close
    Exit Sub
  End If
  
  Winsock1(Index).Close
  m_Ready(Index) = True
    
End Sub

Public Sub DisconnectMe(Index As Integer)
  On Error Resume Next
  Winsock1(Index).SendData "QUIT" & vbCrLf
  DoEvents
  m_Ready(Index) = True
  DoEvents
End Sub

Sub DeleteEmailAccount(ID As Integer)
  DB.Execute ("Delete From EmailAccounts Where ID=" & ID)
End Sub
