VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7110
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   7905
   _ExtentX        =   13944
   _ExtentY        =   12541
   _Version        =   393216
   Description     =   "SmartSearch - Make your Search Smart"
   DisplayName     =   "Smart Search ..."
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mfrmSearch                 As New frmSearch
Public WithEvents MenuHandler As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1

Sub Hide()
    On Error Resume Next
    FormDisplayed = False
    Unload mfrmSearch
End Sub

Sub Show()
    On Error Resume Next
    If mfrmSearch Is Nothing Then
        Set mfrmSearch = New frmSearch
    End If
    Set mfrmSearch.VBInstance = VBInstance
    Set mfrmSearch.Connect = Me
    FormDisplayed = True
    mfrmSearch.Show
End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    Set VBInstance = Application
    If ConnectMode = ext_cm_External Then
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("SmartSearch ...")
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            Me.Show
        End If
    End If
    Exit Sub
error_handler:
    MsgBox Err.Description
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    mcbMenuCommandBar.Delete
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    Unload mfrmSearch
    Set mfrmSearch = Nothing
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        Me.Show
    End If
End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl
    Dim cbMenu As Object
    On Error GoTo AddToAddInCommandBarErr
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        Exit Function
    End If
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    cbMenuCommandBar.FaceId = 2652
    cbMenuCommandBar.Caption = sCaption
    Set AddToAddInCommandBar = cbMenuCommandBar
    Exit Function
AddToAddInCommandBarErr:
End Function

