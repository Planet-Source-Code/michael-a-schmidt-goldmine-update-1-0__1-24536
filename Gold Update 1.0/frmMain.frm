VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   " Gold Update"
   ClientHeight    =   7275
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6735
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   6735
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSimulate 
      Caption         =   "Simulate (Do not write changes to GoldMine)"
      Height          =   225
      Left            =   60
      TabIndex        =   31
      Top             =   6660
      Width           =   3585
   End
   Begin VB.ComboBox cmbDBFields 
      Height          =   315
      Index           =   1
      Left            =   90
      TabIndex        =   28
      Top             =   420
      Width           =   1425
   End
   Begin VB.ComboBox cmbDBFields 
      Height          =   315
      Index           =   2
      Left            =   90
      TabIndex        =   27
      Top             =   780
      Width           =   1425
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process"
      Default         =   -1  'True
      Height          =   345
      Left            =   5490
      TabIndex        =   26
      Top             =   6600
      Width           =   1155
   End
   Begin VB.ComboBox cmbDBFields 
      DataSource      =   "DAO"
      Height          =   315
      Index           =   0
      Left            =   90
      TabIndex        =   14
      Top             =   60
      Width           =   1425
   End
   Begin VB.ComboBox cmbDBFields 
      Height          =   315
      Index           =   12
      Left            =   3360
      TabIndex        =   13
      Top             =   2220
      Width           =   1425
   End
   Begin VB.ComboBox cmbDBFields 
      Height          =   315
      Index           =   11
      Left            =   3360
      TabIndex        =   12
      Top             =   1860
      Width           =   1425
   End
   Begin VB.ComboBox cmbDBFields 
      Height          =   315
      Index           =   10
      Left            =   3360
      TabIndex        =   11
      Top             =   1500
      Width           =   1425
   End
   Begin VB.ComboBox cmbDBFields 
      Height          =   315
      Index           =   9
      Left            =   3360
      TabIndex        =   10
      Top             =   1140
      Width           =   1425
   End
   Begin VB.ComboBox cmbDBFields 
      Height          =   315
      Index           =   8
      Left            =   3360
      TabIndex        =   9
      Top             =   780
      Width           =   1425
   End
   Begin VB.ComboBox cmbDBFields 
      Height          =   315
      Index           =   7
      Left            =   3360
      TabIndex        =   8
      Top             =   420
      Width           =   1425
   End
   Begin VB.ComboBox cmbDBFields 
      Height          =   315
      Index           =   6
      Left            =   90
      TabIndex        =   7
      Top             =   2220
      Width           =   1425
   End
   Begin VB.ComboBox cmbDBFields 
      Height          =   315
      Index           =   5
      Left            =   90
      TabIndex        =   6
      Top             =   1860
      Width           =   1425
   End
   Begin VB.ComboBox cmbDBFields 
      Height          =   315
      Index           =   4
      Left            =   90
      TabIndex        =   5
      Top             =   1500
      Width           =   1425
   End
   Begin VB.ComboBox cmbDBFields 
      Height          =   315
      Index           =   3
      Left            =   90
      TabIndex        =   4
      Top             =   1140
      Width           =   1425
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   6270
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid FlexGrid 
      Bindings        =   "frmMain.frx":0442
      Height          =   2745
      Left            =   60
      TabIndex        =   1
      Top             =   2640
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4842
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16776688
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6150
      Top             =   4740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   7050
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   11377
            Key             =   "goldmine"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLog 
      Height          =   795
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   5430
      Width           =   6585
   End
   Begin VB.Data DAO 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   30
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5850
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Shape Shape13 
      Height          =   315
      Left            =   2010
      Top             =   420
      Width           =   1125
   End
   Begin VB.Line Line13 
      X1              =   1530
      X2              =   2010
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Line Line12 
      X1              =   1530
      X2              =   2010
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Shape Shape12 
      Height          =   315
      Left            =   2010
      Top             =   780
      Width           =   1125
   End
   Begin VB.Label lblGMField 
      AutoSize        =   -1  'True
      Caption         =   "Company"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   2100
      TabIndex        =   30
      Top             =   480
      Width           =   675
   End
   Begin VB.Label lblGMField 
      AutoSize        =   -1  'True
      Caption         =   "Contact"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   2100
      TabIndex        =   29
      Top             =   840
      Width           =   570
   End
   Begin VB.Label lblGMField 
      AutoSize        =   -1  'True
      Caption         =   "Zip"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   5370
      TabIndex        =   25
      Top             =   2280
      Width           =   210
   End
   Begin VB.Label lblGMField 
      AutoSize        =   -1  'True
      Caption         =   "State"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   5370
      TabIndex        =   24
      Top             =   1920
      Width           =   390
   End
   Begin VB.Label lblGMField 
      AutoSize        =   -1  'True
      Caption         =   "City"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   5370
      TabIndex        =   23
      Top             =   1560
      Width           =   285
   End
   Begin VB.Label lblGMField 
      AutoSize        =   -1  'True
      Caption         =   "Address3"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   5370
      TabIndex        =   22
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label lblGMField 
      AutoSize        =   -1  'True
      Caption         =   "Address2"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   5370
      TabIndex        =   21
      Top             =   840
      Width           =   675
   End
   Begin VB.Label lblGMField 
      AutoSize        =   -1  'True
      Caption         =   "Address1"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   5370
      TabIndex        =   20
      Top             =   480
      Width           =   675
   End
   Begin VB.Shape Shape5 
      Height          =   315
      Left            =   5280
      Top             =   780
      Width           =   1125
   End
   Begin VB.Label lblGMField 
      AutoSize        =   -1  'True
      Caption         =   "Fax"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   2100
      TabIndex        =   19
      Top             =   2280
      Width           =   270
   End
   Begin VB.Label lblGMField 
      AutoSize        =   -1  'True
      Caption         =   "Phone3"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   2100
      TabIndex        =   18
      Top             =   1920
      Width           =   540
   End
   Begin VB.Label lblGMField 
      AutoSize        =   -1  'True
      Caption         =   "Phone2"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   2100
      TabIndex        =   17
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label lblGMField 
      AutoSize        =   -1  'True
      Caption         =   "Phone1"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   2100
      TabIndex        =   16
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label lblGMField 
      AutoSize        =   -1  'True
      Caption         =   "AccountNo"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   2970
      TabIndex        =   15
      Top             =   90
      Width           =   780
   End
   Begin VB.Shape Shape11 
      Height          =   315
      Left            =   2790
      Top             =   30
      Width           =   1125
   End
   Begin VB.Line Line11 
      X1              =   1530
      X2              =   2790
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Line Line10 
      X1              =   1530
      X2              =   2010
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Line Line9 
      X1              =   1530
      X2              =   2010
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Line Line8 
      X1              =   1530
      X2              =   2010
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Line Line7 
      X1              =   1530
      X2              =   2010
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Line Line6 
      X1              =   4800
      X2              =   5280
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Line Line5 
      X1              =   4800
      X2              =   5280
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Line Line4 
      X1              =   4800
      X2              =   5280
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Line Line3 
      X1              =   4800
      X2              =   5280
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Line Line2 
      X1              =   4800
      X2              =   5280
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Line Line1 
      X1              =   4800
      X2              =   5280
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Shape Shape10 
      Height          =   315
      Left            =   2010
      Top             =   1140
      Width           =   1125
   End
   Begin VB.Shape Shape9 
      Height          =   315
      Left            =   2010
      Top             =   1500
      Width           =   1125
   End
   Begin VB.Shape Shape8 
      Height          =   315
      Left            =   2010
      Top             =   1860
      Width           =   1125
   End
   Begin VB.Shape Shape7 
      Height          =   315
      Left            =   2010
      Top             =   2220
      Width           =   1125
   End
   Begin VB.Shape Shape6 
      Height          =   315
      Left            =   5280
      Top             =   420
      Width           =   1125
   End
   Begin VB.Shape Shape4 
      Height          =   315
      Left            =   5280
      Top             =   1140
      Width           =   1125
   End
   Begin VB.Shape Shape3 
      Height          =   315
      Left            =   5280
      Top             =   1500
      Width           =   1125
   End
   Begin VB.Shape Shape2 
      Height          =   315
      Left            =   5280
      Top             =   1860
      Width           =   1125
   End
   Begin VB.Shape Shape1 
      Height          =   315
      Left            =   5280
      Top             =   2220
      Width           =   1125
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu smOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu smExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mSettings 
      Caption         =   "Settings"
      Begin VB.Menu smOptions 
         Caption         =   "Options..."
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "Help"
      Begin VB.Menu smAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DBase As Database
Dim GMDBase As Long

Dim DBaseFile As String
Dim DBaseFolder As String

Dim CancelProcess As Boolean

'################################
'   Open DBase
'################################
' Called by cmdOpen_Click which sets our filename and file folder for our database. This sub
' creates a recordset and opens our database, then loops through all our combo-boxes and fills
' each box with the fieldnames in the database.
Private Sub OpenDbase()
Dim Query As String
Dim Counter As Integer
Dim iCount As Integer

    On Error GoTo ErrSub
    Me.MousePointer = vbHourglass
    
    ' [ Open Database - Driver ]
    ' [ Select DBase file as "Table" ]
    Set DBase = OpenDatabase(DBaseFolder, True, False, "dBASE 5.0")
    Query = "SELECT * FROM " & DBaseFile

    Set DAO.Recordset = DBase.OpenRecordset(Query)
    DAO.Refresh
    DBaseOpen = True
    
    ' [ Fill Combo Boxes With Field Names ]
    For iCount = GMFields.Account To GMFields.Zip
        cmbDBFields(iCount).Clear
        For Counter = 0 To DAO.Recordset.Fields.Count - 1
            cmbDBFields(iCount).AddItem Counter & " - " & DAO.Recordset.Fields(Counter).name
            If UCase(DAO.Recordset.Fields(Counter).name) = UCase(lblGMField(iCount).Caption) Then _
                cmbDBFields(iCount).Text = Counter & " - " & DAO.Recordset.Fields(Counter).name
        Next Counter
    Next iCount

    Me.MousePointer = vbNormal
    
Exit Sub

ErrSub:
    Log "Error [" & Err.Number & ":" & Err.Description, "Sub: OpenDBase"
    If MsgBox("Error in OpenDBase - Exit Sub?", vbYesNo) = vbNo Then Resume Next Else Exit Sub
End Sub


'################################
'   Process (Click)
'################################
' This sub opens our GoldMine database, moves to our first local dBase record, and proceeds to loop through
' every dBase record calling an update sub and passing it an account number. The account number is a field
' selected on the form, as with all the other fields.
Private Sub cmdProcess_Click()
On Error GoTo ErrSub

Dim iCount As Integer
Dim iField As Integer

    ' [ Cancel Process ]
    If cmdProcess.Caption = "Cancel" Then
        CancelProcess = True
        Exit Sub
    End If

    ' [ Open GoldMine Database ]
    ' [ Test Data ]
    If Not (GoldMineOpen And DBaseOpen) Then
        MsgBox "GoldMine BDE or DBase not opened!", vbInformation
        Exit Sub
    End If
    ' [ Test Data ]
    If cmbDBFields(GMFields.Account).Text = "" Then
        MsgBox "No Account Relation Selected!", vbInformation
        Exit Sub
    End If
    ' [ Confirm ]
    If chkSimulate.Value = vbUnchecked Then
        If MsgBox("Write changes to GoldMine?", vbYesNo + vbQuestion, "Continue?") = vbNo Then Exit Sub
    End If

    ' [ Disable Form ]
    EnableForm False
    CancelProcess = False

    ' [ Open GoldMine Database ]
    GMDBase = GMW_DB_Open("Contact1")
    If (GMDBase) Then

        ' [ Set Progress Bar Properties ]
        Progress.Max = DAO.Recordset.RecordCount
        Progress.Value = 0
        
        ' [ Create New Log ]
        Open LogName For Output As #LogFile  '-
        Close #LogFile

        ' [ Move To First Record ]
        DAO.Recordset.MoveFirst
        Log Time & " Process Starting."
        
        ' [ Loop Until Complete ]
        For iCount = 1 To DAO.Recordset.RecordCount
            
            ' Update Record - Grab User Defined Field and get index of field for Account...
            iField = Trim(Left(cmbDBFields(GMFields.Account), 2))
            If Not UpdateGMRecord(DAO.Recordset.Fields(iField)) Then
                Log "Error - Update GoldMine record failed!", DAO.Recordset.Fields(iField)
            End If

            ' Update Progress
            Progress.Value = iCount
            StatusBar.Panels(1).Text = "Processed " & iCount & " / " & DAO.Recordset.RecordCount
            
            ' Move Next
            DAO.Recordset.MoveNext
            DoEvents
            
            If CancelProcess = True Then
                Log Time & " Process Canceled."
                GMW_DB_Close (GMDBase)
                EnableForm True
                Exit Sub
            End If

        Next iCount

        ' [ Close and log ]
        DAO.Refresh
        Log Time & " Process Ended."
        GMW_DB_Close (GMDBase)
    Else
        Log "Error - Unable to open GoldMine database!", "Sub: cmdProcess"
    End If

    ' [ Enable Form ]
    EnableForm True
    
Exit Sub

ErrSub:
    If Err.Number = 70 Then
        MsgBox "Could not open Log file, file in use! Please close the log file to continue.", vbCritical, "Error!"
        Open LogName For Output As #LogFile  '-
        Resume Next
    Else
        Log "Error - " & Err.Number & ":" & Err.Description, "Sub: Process"
    End If
End Sub


'################################
'   Update GoldMine Record
'################################
' To update GoldMine, we locate a record by searching for the Acccount Number, which is passed to us in this sub.
' We set our search 'order' which is what field we search by, ie AccountNo. Then we seek the record. Assuming we
' have now located the record, we then proceed to:
'
' Loop through every local (dBase) field, ie combo-box and see if it is 'related' (not empty).
' If the field is 'related' in our loop we update it.
'
' A note on how we know what 'GoldMine' field we are updating. On our form is an array of labels to the right
' of the combo boxes, the text in these fields is passed to the GoldMine API as a GoldMine field name...ie our
' AccountNo label, when we call a function SEEK(FIELD, VALUETOFIND) we pass it the label SEEK(label(X),VAL)...
Public Function UpdateGMRecord(AccountNo As String) As Boolean
Dim CallResult
Dim FailedText
Dim iCount As Integer
Dim iField As Integer
Dim iUpdate As String
        
        ' [ Set Search Index ]
        If GMW_DB_SetOrder(GMDBase, "AccountNo") Then
            Log "Error - Set search index failed!", "Sub: UpdateGMRecord"
            GoTo Failed
        End If
        ' [ Search Record ]
        If GMW_DB_Seek(GMDBase, AccountNo) <> 1 Then
            Log "Seek record failed!"
            GoTo Failed
        End If

        ' [ Found Record ]
        ' Now we loop through each of our fields
        ' and decide if we should update.
        For iCount = GMFields.Account + 1 To GMFields.Zip

            ' If our combo field is not empty, we
            ' must parse the value and update.
            If cmbDBFields(iCount).Text <> "" Then
                
                iField = Trim(Left(cmbDBFields(iCount), 2))   ' Grab local field index.
                iUpdate = ""
                If Not DAO.Recordset.Fields(iField) = vbNull Then iUpdate = DAO.Recordset.Fields(iField)

                ' Read previous entry in goldmine
                ' for logging purposes.
                Dim lName As String
                lName = String$(DAO.Recordset.Fields(iField).Size - 1, Chr(32))
                CallResult = GMW_DB_Read(GMDBase, lblGMField(iCount), lName, DAO.Recordset.Fields(iField).Size)

                ' If both local and goldmine names
                ' are empty, do not log or change.
                lName = Trim(lName)
                If Not (lName = "" And iUpdate = "") Then

                    ' Attempt to change data in GoldMine,
                    ' if fails, ask user to continue?
                    If chkSimulate Then
                        Log lblGMField(iCount) & " - Changed " & lName & " to " & iUpdate, AccountNo
                    ElseIf Not (GMW_DB_Replace(GMDBase, lblGMField(iCount), iUpdate, 0) = 1) Then
                        Log "Could not update GoldMine Field! (Replace Function)"
                        If MsgBox("Resume? (Y/N)", vbCritical + vbYesNo) = vbNo Then GoTo Failed
                    Else
                        Log lblGMField(iCount) & " - Changed " & lName & " to " & iUpdate, AccountNo
                    End If

                End If
            End If
        
        Next iCount

        UpdateGMRecord = True

Exit Function
Failed:
    Log "Update terminated."
    UpdateGMRecord = False
    Exit Function
End Function


Private Sub Form_Load()
    Me.Caption = Me.Caption & " " & App.Major & "." & App.Minor
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
    If Me.WindowState = vbNormal Then
    Me.Height = 7965
    Me.Width = 6855
    End If
End Sub

Private Sub smAbout_Click()
    MsgBox "Gold Update" & vbCrLf & "(C)2001 Michael A. Schmidt", vbInformation, "About..."
End Sub

Private Sub smExit_Click()
    Unload Me
End Sub

Private Sub smOpen_Click()
On Error GoTo ErrSub

  ' Set Dialog Looks / Settings
  CommonDialog.CancelError = True
  CommonDialog.Flags = cdlOFNHideReadOnly
  CommonDialog.Filter = "All Files (*.*)|*.*|dBase Files (*.dbf)|*.dbf"
  CommonDialog.FilterIndex = 2
  CommonDialog.InitDir = App.Path
  CommonDialog.ShowOpen
  DBaseFile = CommonDialog.FileName
  
  ' Parse Returned Path/File
  DBaseFolder = Left(DBaseFile, InStrRev(DBaseFile, "\") - 1)           ' Parse full path...
  DBaseFile = Right(DBaseFile, Len(DBaseFile) - Len(DBaseFolder) - 1)   ' Parse just filename...
  DBaseFile = Left(DBaseFile, Len(DBaseFile) - 4)                       ' Pull ".DBF" off.

  ' Open Database
  Call OpenDbase

Exit Sub
ErrSub:
    DBaseFile = ""
    DBaseFolder = ""
    Log "Error Opening Database! [" & Err.Number & ":" & Err.Description
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If DBaseOpen = True Then DAO.Recordset.Close
    If GoldMineOpen = True Then GMW_UnloadBDE
End Sub

Private Sub smOptions_Click()
    frmSettings.Show vbModal
End Sub


Private Sub EnableForm(YesNo As Boolean)
Dim iCount As Integer

    For iCount = GMFields.Account To GMFields.Zip
        cmbDBFields(iCount).Enabled = YesNo
    Next iCount
    chkSimulate.Enabled = YesNo

    If YesNo = False Then cmdProcess.Caption = "Cancel"
    If YesNo = True Then cmdProcess.Caption = "Process"

End Sub
