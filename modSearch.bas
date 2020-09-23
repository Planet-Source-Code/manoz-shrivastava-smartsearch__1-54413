Attribute VB_Name = "modSearch"
Option Explicit

Public Const LVM_FIRST                      As Long = &H1000
Public Const LVM_SETCOLUMNWIDTH             As Long = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE                 As Long = -1
Public Const LVSCW_AUTOSIZE_USEHEADER       As Long = -2
Private Const SWP_NOACTIVATE        As Long = &H10
Private Const SWP_NOSIZE            As Long = &H1
Private Const SWP_NOMOVE            As Long = &H2
Private Const HWND_TOPMOST          As Long = (-1)
Private Const HWND_NOTOPMOST        As Long = (-2)

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any) As Long

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, _
ByVal X As Long, _
ByVal Y As Long, _
ByVal cx As Long, _
ByVal cy As Long, _
ByVal wFlags As Long)

Public Sub SetTopMost(frm1 As Form, _
    ByVal isTopMost As Boolean)
    
    SetWindowPos frm1.hwnd, IIf(isTopMost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
    DoEvents
    
End Sub
    
Public Function LvAutoSize(lv As ListView)
    Dim Col2Adjst As Long
    Dim LngRc     As Long
    
    For Col2Adjst = 0 To lv.ColumnHeaders.Count - 1
        LngRc = SendMessage(lv.hwnd, LVM_SETCOLUMNWIDTH, Col2Adjst, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next
    
End Function
    
    
