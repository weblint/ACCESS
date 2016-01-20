Attribute VB_Name = "mdlScrollbar"
Option Compare Database
Option Explicit

Private Declare Function apiGetScrollInfo Lib "user32" Alias "GetScrollInfo" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Private Declare Function apiGetWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function apiGetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function apiGetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassname As String, ByVal nMaxCount As Long) As Long

Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_TRACKPOS = &H10

Private Const SB_CTL = 2

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const GWL_STYLE = (-16)

Private Const SBS_VERT = &H1&

Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Public Function fGetScrollBarPos(frm As Form) As Long
' Return ScrollBar Thumb position
' for the Vertical Scrollbar attached to the
' Form passed to this Function.

Dim hWndSB As Long
Dim lngret As Long
Dim sInfo As SCROLLINFO
    
    ' Init SCROLLINFO structure
    sInfo.fMask = SIF_ALL
    sInfo.cbSize = Len(sInfo)
    sInfo.nPos = 0
    sInfo.nTrackPos = 0
    
    ' Call function to get handle to
    ' ScrollBar control if it is visible
    hWndSB = fIsScrollBar(frm)
    If hWndSB = -1 Then
        fGetScrollBarPos = False
        Exit Function
    End If
    
    ' Get the window's ScrollBar position
    lngret = apiGetScrollInfo(hWndSB, SB_CTL, sInfo)
    'Debug.Print "nPos:" & sInfo.nPos & "  nPage:" & sInfo.nPage & "  nMax:" & sInfo.nMax
    'MsgBox "getscrollinfo returned " & sInfo.nPos & " , " & sInfo.nTrackPos
    fGetScrollBarPos = sInfo.nPos + 1

End Function

Private Function fIsScrollBar(frm As Form) As Long
' Get ScrollBar's hWnd
Dim hWnd_VSB As Long
Dim hWnd As Long
   
hWnd = frm.hWnd
    
    ' Let's get first Child Window of the FORM
    hWnd_VSB = apiGetWindow(hWnd, GW_CHILD)
                
    ' Let's walk through every sibling window of the Form
    Do
        ' Thanks to Terry Kreft for explaining
        ' why the apiGetParent acll is not required.
        ' Terry is in a Class by himself! :-)
        'If apiGetParent(hWnd_VSB) <> hWnd Then Exit Do
            
        If fGetClassName(hWnd_VSB) = "NUIscrollBar" Or fGetClassName(hWnd_VSB) = "scrollBar" Then
            If apiGetWindowLong(hWnd_VSB, GWL_STYLE) And SBS_VERT Then
                fIsScrollBar = hWnd_VSB
                Exit Function
            End If
        End If
    
    ' Let's get the NEXT SIBLING Window
    hWnd_VSB = apiGetWindow(hWnd_VSB, GW_HWNDNEXT)
    
    ' Let's Start the process from the Top again
    ' Really just an error check
    Loop While hWnd_VSB <> 0
    
    ' SORRY - NO Vertical ScrollBar control
    ' is currently visible for this Form
    fIsScrollBar = -1
End Function

' From Dev Ashish's Site
' The Access Web
' http://www.mvps.org/access/

'******* Code Start *********
Private Function fGetClassName(hWnd As Long)
Dim strBuffer As String
Dim lngLen As Long
Const MAX_LEN = 255
    strBuffer = Space$(MAX_LEN)
    lngLen = apiGetClassName(hWnd, strBuffer, MAX_LEN)
    If lngLen > 0 Then fGetClassName = Left$(strBuffer, lngLen)
End Function
'******* Code End *********
