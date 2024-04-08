VERSION 5.00
Object = "{2A5B4734-FB68-4DC1-A72D-59DB30D99AFD}#1.2#0"; "ExtraRequestsV2.ocx"
Begin VB.UserControl ExtraReqV2WCtl 
   ClientHeight    =   10005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18990
   KeyPreview      =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   18990
   Begin ExtraRequestsV2.ExtraRequestsCtrlV2 ExtraRequestsCtrlV2 
      Height          =   9735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   17171
   End
End
Attribute VB_Name = "ExtraReqV2WCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Implements LSExtensionWindowLib.IExtensionWindow
Implements LSExtensionWindowLib.IExtensionWindow2

Private NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
Private NtlsSite As LSExtensionWindowLib.IExtensionWindowSite2
Private NtlsUser As LSSERVICEPROVIDERLib.NautilusUser

Private res As New MSXML.DOMDocument
Private Ares As New MSXML.DOMDocument
Private ProcessXML As LSSERVICEPROVIDERLib.NautilusProcessXML

Private sp As LSSERVICEPROVIDERLib.NautilusServiceProvider



Private con As ADODB.Connection


Public Function IExtensionWindow_CloseQuery() As Boolean

10    IExtensionWindow_CloseQuery = True

End Function

Public Function IExtensionWindow_DataChange() As LSExtensionWindowLib.WindowRefreshType
10    IExtensionWindow_DataChange = windowRefreshNone
End Function

Private Function IExtensionWindow_GetButtons() As LSExtensionWindowLib.WindowButtonsType
10        IExtensionWindow_GetButtons = windowButtonsNone
End Function

Public Sub IExtensionWindow_Internationalise()

End Sub


Public Sub IExtensionWindow_PreDisplay()
10    On Error GoTo ERR_ashi
          
        
'          Dim constr As String
'
'
'20        Set con = New ADODB.Connection
'30        Call con.Open(NtlsCon.GetADOConnectionString)
'40        con.CursorLocation = adUseClient
'
'50        con.Execute "SET ROLE LIMS_USER"
'60        Call ConnectSameSession(CDbl(NtlsCon.GetSessionId))


          
70        Exit Sub
ERR_ashi:
80     MsgBox " ERR_ashi : " & Err.Description
       
End Sub

Public Sub IExtensionWindow_refresh()
'code for refreshing the window
'    Call RefreshWindow
End Sub

Public Sub IExtensionWindow_RestoreSettings(ByVal hKey As Long)

End Sub

Public Function IExtensionWindow_SaveData() As Boolean

End Function

Public Sub IExtensionWindow_SaveSettings(ByVal hKey As Long)

End Sub

Public Sub IExtensionWindow_SetParameters(ByVal parameters As String)
10       On Error GoTo ERR_IExtensionWindow_SetParameters
           

20        Exit Sub
ERR_IExtensionWindow_SetParameters:
30        MsgBox "ERR_IExtensionWindow_SetParameters" & vbCrLf & Err.Description
End Sub

Public Sub IExtensionWindow_SetServiceProvider(ByVal ServiceProvider As Object)
10        On Error GoTo ErrHnd
          
20        Set sp = ServiceProvider
30        Set ProcessXML = sp.QueryServiceProvider("ProcessXML")
40        Set NtlsCon = sp.QueryServiceProvider("DBConnection")
50        Set NtlsUser = sp.QueryServiceProvider("User")
          
60        Exit Sub
ErrHnd:
70        MsgBox "IExtensionWindow_SetServiceProvider"
End Sub

Public Sub IExtensionWindow_SetSite(ByVal Site As Object)
10       On Error GoTo ErrHnd
           

20        Set NtlsSite = Site

30        NtlsSite.SetWindowInternalName ("ExtraRequestsCtrlV2")
40        NtlsSite.SetWindowRegistryName ("ExtraRequestsCtrlV2")
               
50        Exit Sub
ErrHnd:
60        MsgBox "IExtensionWindow_SetSite"
End Sub

Public Sub IExtensionWindow_Setup()
              
                   


          
           
10         ExtraRequestsCtrlV2.RunFromWindow = True
20         Call ExtraRequestsCtrlV2.IExtensionWindow_SetServiceProvider(sp)
30         ExtraRequestsCtrlV2.IExtensionWindow_Internationalise
           
40         ExtraRequestsCtrlV2.IExtensionWindow_PreDisplay
50         ExtraRequestsCtrlV2.IExtensionWindow_GetButtons
60         ExtraRequestsCtrlV2.IExtensionWindow_Setup


End Sub
Private Sub ExtraRequestsCtrlV2_CloseClicked()

10        If MsgBox(" ? האם אתה בטוח שברצונך לצאת ", vbYesNo + vbDefaultButton2) = vbNo Then
              
20                Exit Sub
30            End If
40            ExtraRequestsCtrlV2.IExtensionWindow_CloseQuery
50            NtlsSite.CloseWindow
End Sub

Private Sub CytoGeneticQuestCtrl_CloseClicked()
        
10          If MsgBox(" ? האם אתה בטוח שברצונך לצאת ", vbYesNo + vbDefaultButton2) = vbNo Then
20                Exit Sub
30            End If
              
40                       ExtraRequestsCtrlV2.IExtensionWindow_CloseQuery
50                       NtlsSite.CloseWindow
                     Set con = Nothing
        
End Sub

Public Function IExtensionWindow_ViewRefresh() As LSExtensionWindowLib.WindowRefreshType
10        IExtensionWindow_ViewRefresh = windowRefreshNone
End Function

Private Sub IExtensionWindow2_Close()
End Sub

Private Sub ConnectSameSession(ByVal aSessionID)
10        On Error GoTo ErrHnd
          Dim aProc As New ADODB.Command
          Dim aSession As New ADODB.Parameter
          
20        aProc.ActiveConnection = con
30        aProc.CommandText = "lims.lims_env.connect_same_session"
40        aProc.CommandType = adCmdStoredProc

50        aSession.Type = adDouble
60        aSession.Direction = adParamInput
70        aSession.Value = aSessionID
80        aProc.parameters.Append aSession

90        aProc.Execute
100       Set aSession = Nothing
110       Set aProc = Nothing
120       Exit Sub
ErrHnd:
130      MsgBox "ConnectSameSession" & vbCrLf & Err.Description
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
          Dim strVer As String
10        On Error GoTo Err_UserControl_KeyDown

20           Exit Sub
Err_UserControl_KeyDown:
30    MsgBox "Err_UserControl_KeyDown : " & Err.Description
End Sub



