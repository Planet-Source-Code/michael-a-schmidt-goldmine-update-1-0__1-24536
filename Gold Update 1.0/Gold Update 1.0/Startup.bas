Attribute VB_Name = "Startup"
'################################
'   Startup (SubMain)
'################################
Public Software As Software_Object  ' Stores all our GoldMine Settings

Public DBaseOpen As Boolean         ' Is local database open?
Public GoldMineOpen As Boolean      ' Is GoldMine BDE open?

Public LogFile As Long             ' Log File Variables
Public LogName As String

Type Software_Object
    GoldMineROOT   As String
    GoldMineBASE   As String
    GoldMineCOMMON As String
    GoldMineUSER   As String
    GoldMinePASS   As String
End Type

' [ GMFields ]
' We use these in FOR loops, these correspond also to our frmMain combo arrays and label arrays. IE you
' will see lblGMField(0) is set a value of 'Account'. This makes adding another GoldMine Contact1 field
' a little easier, just make sure these values (Account = Index 0) correspond with the frmMain labels
' and combo-boxes, also check any FOR loops in the program, so if you add "DEAR" after Zip, you would
' change the loop
' For X = GMFields.Account to GMFields.Zip   To   For X = GMFields.Account to GMFields.Dear
Public Enum GMFields
    Account
    Company
    Contact
    Phone1
    Phone2
    Phone3
    Fax
    Address1
    Address2
    Address3
    City
    State
    Zip
End Enum


'################################
'   Sub Main
'################################
Sub Main()
On Error GoTo ErrSub

    ' [ Load Log File ]
    LogName = App.Path & "\Log.csv"
    LogFile = FreeFile()
    Open LogName For Output As #LogFile  '-
    Close #LogFile

    ' [ Load All Registry Settings ]
    LoadAllSettings

    ' [ Load our main form so we can use the log ]
    GoldMineOpen = False
    DBaseOpen = False
    frmMain.Show

    ' [ Load GoldMine API ]
    If LoadGoldMineAPI Then frmMain.StatusBar.Panels(1).Text = "GoldMine BDE Loaded." Else _
                            frmMain.StatusBar.Panels(1).Text = "GoldMine BDE Failed!"

Exit Sub
ErrSub:
    If Err.Number = 70 Then
        MsgBox "Could not open Log file, file in use!", vbCritical, "Error!"
        End
    End If
End Sub


'################################
'   Registry Functions
'################################
Public Sub ResetAllSettings()
On Error Resume Next
    DeleteSetting (App.Title)
    LoadAllSettings

End Sub
Public Sub LoadAllSettings()

    ' Load Settings
    Software.GoldMineBASE = GetSetting(App.Title, "PATH", "GMBASE", "\\kingkong\Apps\GoldMine\gmbase\")
    Software.GoldMineCOMMON = GetSetting(App.Title, "PATH", "GMCOMMON", "\\kingkong\Apps\GoldMine\Common\")
    Software.GoldMineROOT = GetSetting(App.Title, "PATH", "GMROOT", "\\kingkong\Apps\GoldMine\")
    Software.GoldMineUSER = GetSetting(App.Title, "PATH", "GMUSER", "MASTER")
    Software.GoldMinePASS = GetSetting(App.Title, "PATH", "GMPASS", "")

End Sub
Public Sub SaveAllSettings()
    
    ' Save Settings
    Call SaveSetting(App.Title, "PATH", "GMBASE", Software.GoldMineBASE)
    Call SaveSetting(App.Title, "PATH", "GMCOMMON", Software.GoldMineCOMMON)
    Call SaveSetting(App.Title, "PATH", "GMROOT", Software.GoldMineROOT)
    Call SaveSetting(App.Title, "PATH", "GMUSER", Software.GoldMineUSER)
    Call SaveSetting(App.Title, "PATH", "GMPASS", Software.GoldMinePASS)
End Sub


'################################
'   Load GoldMine API (BDE)
'################################
Public Function LoadGoldMineAPI() As Boolean
Dim CallResult As Long

    LoadGoldMineAPI = False

    CallResult = GMW_LoadBDE(Software.GoldMineROOT, Software.GoldMineBASE, Software.GoldMineCOMMON, Software.GoldMineUSER, Software.GoldMinePASS)
    If CallResult <> 1 Then
        Select Case CallResult
        Case 0: Log "BDE Error: Already Loaded!"
        Case -1: Log "BDE Error: Failed to load!"
        Case -2: Log "BDE Error: Cannot find license file!"
        Case -3: Log "BDE Error: Cannot load license file!"
        Case -4: Log "BDE Error: Cannot validate the license file!"
        End Select

        GoldMineOpen = False
        LoadGoldMineAPI = False
        Exit Function
    End If
    
    GoldMineOpen = True
    LoadGoldMineAPI = True

End Function


'################################
'   Log
'################################
Public Sub Log(Optional Field1 As String, Optional Field2 As String, Optional Field3 As String)
Dim CVSString As String
Dim LOGString As String
On Error GoTo ErrSub

    CVSString = Time & ",""" & Field1 & """,""" & Field2 & """,""" & Field3 & """"
    LOGString = Trim("(" & Time & ") " & Field1 & " " & Field2 & " " & Field3)

    frmMain.txtLog.Text = LOGString & vbCrLf & frmMain.txtLog.Text
    If Len(frmMain.txtLog.Text) > 1000 Then frmMain.txtLog.Text = Left(frmMain.txtLog.Text, 1000)

    ' Print to LOG
    Open LogName For Append As #LogFile
    Print #LogFile, CVSString
    Close #LogFile

Exit Sub
ErrSub:
    If Err.Number = 70 Then
        MsgBox "Could not open Log file, file in use! Please close the log file to continue.", vbCritical, "Error!"
        Open LogName For Output As #LogFile  '-
        Resume Next
    Else
        Log "Error - " & Err.Number & ":" & Err.Description, "Sub: Log"
    End If
End Sub
