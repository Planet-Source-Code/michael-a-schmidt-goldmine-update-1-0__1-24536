Attribute VB_Name = "GoldMineAPI"
'################################
'   GoldMine API
'################################


Public Type GMLicInfo
    Licensee As String * 60
    LicNo As String * 20
    SiteName As String * 20
    LicUsers As Long
    SQLUsers As Long
    GSSites As Long
    IsDemo As Long
    IsServerLic As Long
    IsRemoteLic As Long
    ISUSALic As Long
    iReserved1 As Long
    iReserved2 As Long
    iReserved3 As Long
    sReserved As String * 100
End Type


' LoadBDE Functions
Public Declare Function GMW_LoadBDE Lib "GM5S32.dll" (ByVal sSysDir As String, ByVal sGoldDir As String, ByVal sCommonDir As String, ByVal sUser As String, ByVal sPassword As String) As Long
Public Declare Function GMW_UnloadBDE Lib "GM5S32.dll" () As Long
Public Declare Function GMW_SetSQLUserPass Lib "GM5S32.dll" (ByVal sUserName As String, ByVal sPassword As String) As Long
' Business logic functions
' Name-Value parameter passing to business logic function
Public Declare Function GMW_Execute Lib "GM5S32.dll" (ByVal szFunc As String, ByVal GMPtr As Any) As Long
Public Declare Function GMW_NV_Create Lib "GM5S32.dll" () As Long
Public Declare Function GMW_NV_CreateCopy Lib "GM5S32.dll" (ByVal hgmnv As Long) As Long
Public Declare Function GMW_NV_Delete Lib "GM5S32.dll" (ByVal hgmnv As Long) As Long
Public Declare Function GMW_NV_Copy Lib "GM5S32.dll" (ByVal hgmnvDestination As Long, ByVal hgmnvSource As Long) As Long
Public Declare Function GMW_GetLicenseInfo Lib "GM5S32.dll" (ByRef LicInfo As Any) As Long
Public Declare Function GMW_NV_GetValue Lib "GM5S32.dll" (ByVal hgmnv As Long, ByVal name As String, ByVal DefaultValue As String) As Long
Public Declare Function GMW_NV_SetValue Lib "GM5S32.dll" (ByVal hgmnv As Long, ByVal name As String, ByVal Value As String) As Long
Public Declare Function GMW_NV_NameExists Lib "GM5S32.dll" (ByVal hgmnv As Long, ByVal name As String) As Long
Public Declare Function GMW_NV_EraseName Lib "GM5S32.dll" (ByVal hgmnv As Long, ByVal name As String) As Long
Public Declare Function GMW_NV_EraseAll Lib "GM5S32.dll" (ByVal hgmnv As Long) As Long
Public Declare Function GMW_NV_Count Lib "GM5S32.dll" (ByVal hgmnv As Long) As Long
Public Declare Function GMW_NV_GetNameFromIndex Lib "GM5S32.dll" (ByVal hgmnv As Long, ByVal index As Long) As Long
Public Declare Function GMW_NV_GetValueFromIndex Lib "GM5S32.dll" (ByVal hgmnv As Long, ByVal index As Long) As Long
' Low-Level DB funcs
Public Declare Function GMW_DB_Open Lib "GM5S32.dll" (ByVal sTableName As String) As Long
Public Declare Function GMW_DB_Close Lib "GM5S32.dll" (ByVal lArea As Long) As Long
Public Declare Function GMW_DB_Append Lib "GM5S32.dll" (ByVal lArea As Long, ByVal sRecID As String) As Long
Public Declare Function GMW_DB_Replace Lib "GM5S32.dll" (ByVal lArea As Long, ByVal sField As String, ByVal sData As String, ByVal iAddTo As Long) As Long
Public Declare Function GMW_DB_Delete Lib "GM5S32.dll" (ByVal lArea As Long) As Long
Public Declare Function GMW_DB_UnLock Lib "GM5S32.dll" (ByVal lArea As Long) As Long
Public Declare Function GMW_DB_Read Lib "GM5S32.dll" (ByVal lArea As Long, ByVal sField As String, ByVal sbuf As String, ByVal lbufsize As Long) As Long
Public Declare Function GMW_DB_Top Lib "GM5S32.dll" (ByVal lArea As Long) As Long
Public Declare Function GMW_DB_Bottom Lib "GM5S32.dll" (ByVal lArea As Long) As Long
Public Declare Function GMW_DB_SetOrder Lib "GM5S32.dll" (ByVal lArea As Long, ByVal Stag As String) As Long
Public Declare Function GMW_DB_Seek Lib "GM5S32.dll" (ByVal lArea As Long, ByVal sParam As String) As Long
Public Declare Function GMW_DB_Skip Lib "GM5S32.dll" (ByVal lArea As Long, ByVal lSkip As Long) As Long
Public Declare Function GMW_DB_Goto Lib "GM5S32.dll" (ByVal lArea As Long, ByVal sRecNo As String) As Long
Public Declare Function GMW_DB_Move Lib "GM5S32.dll" (ByVal lArea As Long, ByVal sCommand As String, ByVal sParam As String) As Long
Public Declare Function GMW_DB_Search Lib "GM5S32.dll" (ByVal lArea As Long, ByVal sExpr As String, ByVal sRecID As String) As Long
Public Declare Function GMW_DB_Filter Lib "GM5S32.dll" (ByVal lArea As Long, ByVal sFilterExpr As String) As Long
Public Declare Function GMW_DB_Range Lib "GM5S32.dll" (ByVal lArea As Long, ByVal sMin As String, ByVal sMax As String, ByVal Stag As String) As Long
Public Declare Function GMW_DB_RecNo Lib "GM5S32.dll" (ByVal lArea As Long, ByVal sRecID As String) As Long
Public Declare Function GMW_DB_IsSQL Lib "GM5S32.dll" (ByVal lArea As Long) As Long
Public Declare Function GMW_NewRecID Lib "GM5S32.dll" (ByVal szRecid As String, ByVal szUser As String) As String
Public Declare Function GMW_UpdateSyncLog Lib "GM5S32.dll" (ByVal sTable As String, ByVal sRecID As String, ByVal sField As String, byvalsAction As String) As Long
Public Declare Function GMW_ReadImpTLog Lib "GM5S32.dll" (ByVal sFile As String, lDelWhenDone As Long, sStatus As String) As Long
Public Declare Function GMW_SyncStamp Lib "GM5S32.dll" (sStamp As String, sOutBuf As String) As Long
' Datastream funcs
Public Declare Function GMW_DS_Query Lib "GM5S32.dll" (ByVal sSQL As String, ByVal sFilter As String, ByVal sFDlm As String, ByVal sRDlm As String) As Long
Public Declare Function GMW_DS_Range Lib "GM5S32.dll" (ByVal sTable As String, ByVal Stag As String, ByVal sTopLimit As String, ByVal sBotLimit As String, ByVal sFields As String, ByVal sFilter As String, ByVal sFDlm As String, ByVal sRDlm As String) As Long
Public Declare Function GMW_DS_Fetch Lib "GM5S32.dll" (ByVal iHandle As Long, ByVal sbuf As String, ByVal iBufSize As Long, ByVal iGetRecs As Long) As Long
Public Declare Function GMW_DS_Close Lib "GM5S32.dll" (ByVal iHandle As Long) As Long
Public Declare Function GMW_IsUserGroupMember Lib "GM5S32.dll" (ByVal szGroup As String, ByVal szUserID As String) As Long
' Misc WinAPI funcs used by VB with the GM5S32.DLL
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
' NOTE! All GM5S32 Funcs that return a string pointer should be
' converted using the following function. For example:
'
' sResult = PtrToStr(GMW_NV_GetValue(lGMPtr, "OutPut", ""))
Public Function PtrToStr(ByVal lpsz As Long) As String
Dim strOut As String
Dim lngStrLen As Long

    lngStrLen = lstrlen(ByVal lpsz)
    ' If returning larger packets, you may have to
    ' increase this value
    lngStrLen = 64000
    
    If (lngStrLen > 0) Then
        strOut = String$(lngStrLen, vbNullChar)
        Call CopyMemory(ByVal strOut, ByVal lpsz, lngStrLen)
        lngStrLen = lstrlen(strOut)
        PtrToStr = Left(strOut, lngStrLen)
    Else
        PtrToStr = ""
    End If
    
    strOut = ""
End Function
