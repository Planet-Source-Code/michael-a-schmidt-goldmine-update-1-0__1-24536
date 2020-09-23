VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5160
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Reset"
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Top             =   2340
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1650
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   3870
      TabIndex        =   7
      Top             =   2340
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   2640
      TabIndex        =   6
      Top             =   2340
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   90
      Width           =   5025
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2100
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   270
         Width           =   1185
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   150
         TabIndex        =   0
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "- User"
         Height          =   210
         Left            =   1590
         TabIndex        =   17
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "- Password"
         Height          =   210
         Left            =   3360
         TabIndex        =   16
         Top             =   300
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   90
      TabIndex        =   8
      Top             =   840
      Width           =   5025
      Begin VB.CommandButton cmdContact1 
         Caption         =   ".."
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   990
         Width           =   375
      End
      Begin VB.CommandButton cmdCal 
         Caption         =   ".."
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   630
         Width           =   375
      End
      Begin VB.CommandButton cmdLicense 
         Caption         =   ".."
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   375
      End
      Begin VB.TextBox txtCommon 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   540
         TabIndex        =   4
         Top             =   990
         Width           =   3195
      End
      Begin VB.TextBox txtGMBase 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   540
         TabIndex        =   3
         Top             =   630
         Width           =   3195
      End
      Begin VB.TextBox txtSysDir 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   540
         TabIndex        =   2
         Top             =   300
         Width           =   3195
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "- Contact1.dbf"
         Height          =   195
         Left            =   3810
         TabIndex        =   11
         Top             =   1050
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "- Cal.dbf"
         Height          =   195
         Left            =   3810
         TabIndex        =   10
         Top             =   690
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "- License.dbf"
         Height          =   195
         Left            =   3810
         TabIndex        =   9
         Top             =   300
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCal_Click()
On Error GoTo ErrSub

  ' Set Dialog Looks / Settings
  CommonDialog.CancelError = True
  CommonDialog.Filter = "Cal File (Cal.dbf)|Cal.dbf|dBase Files (*.dbf)|*.dbf|All Files (*.*)|*.*"
  CommonDialog.FileName = "Cal.dbf"
  CommonDialog.ShowOpen
    txtGMBase = Left(CommonDialog.FileName, InStrRev(CommonDialog.FileName, "\"))           ' Parse full path...

Exit Sub
ErrSub:
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdContact1_Click()
On Error GoTo ErrSub

  ' Set Dialog Looks / Settings
  CommonDialog.CancelError = True
  CommonDialog.Filter = "Contact File (Contact1.dbf)|Contact1.dbf|dBase Files (*.dbf)|*.dbf|All Files (*.*)|*.*"
  CommonDialog.FileName = "Contact1.dbf"
  CommonDialog.ShowOpen
    txtCommon = Left(CommonDialog.FileName, InStrRev(CommonDialog.FileName, "\"))           ' Parse full path...

Exit Sub
ErrSub:
    Exit Sub
End Sub

Private Sub cmdLicense_Click()
On Error GoTo ErrSub

  ' Set Dialog Looks / Settings
  CommonDialog.CancelError = True
  CommonDialog.Filter = "License File (License.dbf)|License.dbf|dBase Files (*.dbf)|*.dbf|All Files (*.*)|*.*"
  CommonDialog.FileName = "License.dbf"
  CommonDialog.ShowOpen
    txtSysDir = Left(CommonDialog.FileName, InStrRev(CommonDialog.FileName, "\"))           ' Parse full path...


Exit Sub
ErrSub:
    Exit Sub
End Sub



Private Sub cmdSave_Click()

    cmdSave.Enabled = False
    Me.MousePointer = vbHourglass
    
    Software.GoldMineUSER = txtUser
    Software.GoldMinePASS = txtPassword
    Software.GoldMineROOT = txtSysDir
    Software.GoldMineBASE = txtGMBase
    Software.GoldMineCOMMON = txtCommon
    
    SaveAllSettings
    
    If LoadGoldMineAPI Then frmMain.StatusBar.Panels(1).Text = "GoldMine BDE Loaded." Else _
                            frmMain.StatusBar.Panels(1).Text = "GoldMine BDE Failed!"
    
    Me.MousePointer = vbNormal
    Me.Enabled = True
    Unload Me
End Sub

Private Sub Command1_Click()
    ResetAllSettings
    FillFormFields
End Sub

Private Sub Form_Load()
    FillFormFields
End Sub


Private Sub FillFormFields()
    txtUser = Software.GoldMineUSER
    txtPassword = Software.GoldMinePASS
    txtSysDir = Software.GoldMineROOT
    txtGMBase = Software.GoldMineBASE
    txtCommon = Software.GoldMineCOMMON
End Sub


'##################################
'   Misc Eye Candy (Focus)
'##################################
Private Sub txtuser_LostFocus()
    txtUser.BackColor = &H80000005
End Sub
Private Sub txtpassword_LostFocus()
    txtPassword.BackColor = &H80000005
End Sub
Private Sub txtsysdir_LostFocus()
    txtSysDir.BackColor = &H80000005
End Sub
Private Sub txtgmbase_LostFocus()
    txtGMBase.BackColor = &H80000005
End Sub
Private Sub txtcommon_LostFocus()
    txtCommon.BackColor = &H80000005
End Sub
Private Sub txtuser_GotFocus()
    txtUser.BackColor = &H80000018
End Sub
Private Sub txtpassword_GotFocus()
    txtPassword.BackColor = &H80000018
End Sub
Private Sub txtsysdir_GotFocus()
    txtSysDir.BackColor = &H80000018
End Sub
Private Sub txtgmbase_GotFocus()
    txtGMBase.BackColor = &H80000018
End Sub
Private Sub txtcommon_GotFocus()
    txtCommon.BackColor = &H80000018
End Sub
