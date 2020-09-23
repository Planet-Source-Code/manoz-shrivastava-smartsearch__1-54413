VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   Caption         =   "SmartSearch"
   ClientHeight    =   6210
   ClientLeft      =   2190
   ClientTop       =   1950
   ClientWidth     =   6825
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwProc 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4471
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483646
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtFind 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   285
      Left            =   6000
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin MSComctlLib.ListView lvwCode 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Component"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Procedure"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Line #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Matching Line"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance   As VBIDE.VBE
Public Connect      As Connect
Dim VBProj          As VBProject
Dim VBComp          As VBComponent
Dim lvw             As ListItem

Option Explicit

Private Sub CancelButton_Click()
    Connect.Hide
End Sub
    
Private Sub Form_Load()
    Dim startLine       As Long
    Dim startCol        As Long
    Dim endLine         As Long
    Dim endCol          As Long
    On Error Resume Next
    
    Caption = Caption & " - [" & VBInstance.ActiveVBProject.Name & "]"
    With VBInstance.ActiveCodePane
        .GetSelection startLine, startCol, endLine, endCol
        If endCol = 1 Then endLine = endLine - 1
        txtFind.Text = Mid$(.CodeModule.Lines(startLine, 1), startCol, endCol - startCol)
    End With
    With lvwProc
        .ColumnHeaders.Add , , " "
        .View = lvwIcon
    End With
    If Len(txtFind.Text) Then
        OKButton_Click
    End If
End Sub
    
Private Sub Form_Unload(Cancel As Integer)
    CancelButton_Click
End Sub
    
Private Sub lvwCode_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim startLine       As Long
    Dim startCol        As Long
    Dim endLine         As Long
    Dim endCol          As Long
    Dim strProcedure    As String
    Dim strComponent    As String
    Dim i               As Long
    Dim lvwDef          As ListItem
    Dim strCode         As String
    
    lvwProc.ListItems.Clear
    lvwProc.View = lvwReport
    strProcedure = Item.SubItems(1)
    strComponent = Item.Tag
    startLine = CInt(Item.SubItems(2))
    strCode = Item.SubItems(3)
    VBInstance.ActiveVBProject.VBComponents(strComponent).CodeModule.CodePane.Show
    With VBInstance.ActiveCodePane.CodeModule
        If .Find(strCode, startLine, startCol, endLine, endCol) Then
            .VBE.ActiveCodePane.SetSelection startLine, startCol, endLine, endCol
        End If
        lvwProc.ColumnHeaders(1).Text = .VBE.SelectedVBComponent.Name & "." & strProcedure
        For i = startLine To .CountOfLines
            Set lvwDef = lvwProc.ListItems.Add(, , .Lines(i, 1))
            If InStr(.Lines(i, 1), "End Sub") > 0 Or _
            InStr(.Lines(i, 1), "End Function") > 0 Or _
            InStr(.Lines(i, 1), "End Property") > 0 Then
            Exit For
        End If
    Next i
End With
Call LvAutoSize(lvwProc)
End Sub
    
Private Sub OKButton_Click()
    Dim strFound        As String
    Dim strCode         As String
    Dim i               As Long
    Dim CModule         As CodeModule
    Dim startLine       As Long
    Dim startCol        As Long
    Dim endLine         As Long
    Dim endCol          As Long
    Dim strProcName     As String
    
    On Error Resume Next
    
    strFound = Trim(txtFind.Text)
    lvwCode.ListItems.Clear
    For Each VBProj In VBInstance.VBProjects
        On Error Resume Next
        For Each VBComp In VBProj.VBComponents
            On Error Resume Next
            With VBComp
                Set CModule = .CodeModule
            End With
            startLine = 1
            Do
                DoEvents
                startCol = 1
                endLine = -1
                endCol = -1
                If CModule.Find(strFound, startLine, startCol, endLine, endCol) Then
                    DoEvents
                    strProcName = GetProcedureName(CModule, startLine)
                    strCode = CModule.Lines(startLine, 1)
                    Set lvw = lvwCode.ListItems.Add(, , CModule.Parent.Name & " - " & GetComponentTypeName(CModule))
                    lvw.SubItems(1) = strProcName
                    lvw.SubItems(2) = CStr(startLine)
                    lvw.SubItems(3) = strCode
                    lvw.Tag = CModule.Parent.Name
                End If
                startLine = startLine + 1
                If startLine >= CModule.CountOfLines Then
                    Exit Do
                End If
            Loop While CModule.Find(strFound, startLine, 1, -1, -1)
        Next VBComp
    Next VBProj
    lvwCode.ColumnHeaders(4).Text = "Matching Line(s) [" & lvwCode.ListItems.Count & "] found."
    Call LvAutoSize(lvwCode)
End Sub
    
Public Function GetProcedureName(CMod As CodeModule, _
    ByVal Sline As Long) As String
    
    On Error Resume Next
    GetProcedureName = CMod.ProcOfLine(Sline, vbext_pk_Proc)
    If LenB(GetProcedureName) = 0 Then
        GetProcedureName = CMod.ProcOfLine(Sline, vbext_pk_Let)
    End If
    If LenB(GetProcedureName) = 0 Then
        GetProcedureName = CMod.ProcOfLine(Sline, vbext_pk_Get)
    End If
    If LenB(GetProcedureName) = 0 Then
        GetProcedureName = CMod.ProcOfLine(Sline, vbext_pk_Set)
    End If
    If LenB(GetProcedureName) = 0 Then
        
        GetProcedureName = "(Declarations)"
    End If
    On Error GoTo 0
    
End Function
    
Private Sub txtFind_GotFocus()
    With txtFind
        .SelStart = 0
        .SelLength = Len(txtFind)
    End With
    lvwProc.ListItems.Clear
End Sub
    
Public Function GetComponentTypeName(codeMod As CodeModule) As String
    Select Case codeMod.Parent.Type
        Case vbext_ct_StdModule
            GetComponentTypeName = " [Bas Module]"
        Case vbext_ct_ClassModule
            GetComponentTypeName = " [Class Module]"
        Case vbext_ct_MSForm
            GetComponentTypeName = " [Form]"
        Case vbext_ct_ResFile
            GetComponentTypeName = " [Resource File]"
        Case vbext_ct_VBForm
            GetComponentTypeName = " [VB Form]"
        Case vbext_ct_VBMDIForm
            GetComponentTypeName = " [MDIForm]"
        Case vbext_ct_PropPage
            GetComponentTypeName = " [PropertyPage]"
        Case vbext_ct_UserControl
            GetComponentTypeName = " [UserControl]"
        Case vbext_ct_DocObject
            GetComponentTypeName = " [RelatedDocument]"
        Case vbext_ct_ActiveXDesigner
            GetComponentTypeName = " [ActiveXDesigner]"
    End Select
End Function
    
    
