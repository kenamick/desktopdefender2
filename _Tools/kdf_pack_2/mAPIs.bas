Attribute VB_Name = "mAPIs"
Option Explicit

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetOpenFileName Lib "COMDLG32" _
    Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "COMDLG32" _
    Alias "GetSaveFileNameA" (file As OPENFILENAME) As Long
Public Declare Function GetFileTitle Lib "COMDLG32" _
    Alias "GetFileTitleA" (ByVal szFile As String, _
    ByVal szTitle As String, ByVal cbBuf As Long) As Long

Declare Function LoadImage Lib "user32" Alias "LoadImageA" ( _
    ByVal hInst As Long, _
    ByVal lpsz As String, _
    ByVal un1 As Long, _
    ByVal n1 As Long, ByVal n2 As Long, _
    ByVal un2 As Long _
    ) As Long

Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function ChooseColor Lib "COMDLG32.DLL" _
    Alias "ChooseColorA" (Color As TCHOOSECOLOR) As Long

Public Type TCHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Public Enum EChooseColor
    CC_RGBINIT = &H1
    CC_FULLOPEN = &H2
    CC_PreventFullOpen = &H4
    CC_ColorShowHelp = &H8
' Win95 only
    CC_SOLIDCOLOR = &H80
    CC_AnyColor = &H100
' End Win95 only
    CC_ENABLEHOOK = &H10
    CC_ENABLETEMPLATE = &H20
    CC_EnableTemplateHandle = &H40
End Enum

Public Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hWndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Public Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

Public Const MAX_PATH = 260
Public Const MAX_FILE = 260
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1
Public Const IMAGE_CURSOR = 2
Public Const LR_COLOR = &H2
Public Const LR_COPYDELETEORG = &H8
Public Const LR_COPYFROMRESOURCE = &H4000
Public Const LR_COPYRETURNORG = &H4
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_DEFAULTSIZE = &H40
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_MONOCHROME = &H1
Public Const LR_SHARED = &H8000

Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCINVERT = &H660046   ' (DWORD) dest = source XOR dest
Public Const SM_CYSCREEN = 1
Public Const SM_CXSCREEN = 0


Public Sub OpenFile(lpFileName As String, lpFilter As String, _
                    Optional DefFileFilter As String, _
                    Optional bOpenFile As Boolean, _
                    Optional strTitle As String = "")
 
 Dim of             As OPENFILENAME
 Dim rval           As Long
 Dim buffer         As String * 260
 Dim s              As String
 Dim ch             As String
 Dim cn             As Integer
 Static strLastDir  As String
 
 buffer = String(Len(buffer), 0)
 
 With of
   .lStructSize = Len(of)
   .hWndOwner = frmMain.hwnd                        ' set window owner
   
   For cn = 1 To Len(lpFilter)                      ' format filter
        ch = Mid$(lpFilter, cn, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
   
   If (Len(strTitle) > 0) Then
    .lpstrTitle = strTitle
   End If
   .lpstrFilter = s & vbNullChar & vbNullChar
   .lpstrFile = buffer
   .nMaxFile = MAX_FILE
   .lpstrFileTitle = buffer
   .nMaxFileTitle = MAX_PATH
   '.lpstrInitialDir = App.Path                      ' set initial dir
   .lpstrInitialDir = strLastDir
   .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY
 
  If (bOpenFile) Then
     rval = GetOpenFileName(of)
  Else
     rval = GetSaveFileName(of)
  End If
  
  If rval = 1 Then                                   ' API was successful
     lpFileName = .lpstrFile
   
   If bOpenFile Then                                 ' do some adjustments on 'open file'
     For cn = 1 To Len(lpFileName)                   ' discard 0s
      If Mid$(lpFileName, cn, 1) = vbNullChar Then Mid$(lpFileName, cn, 1) = " "
     Next
     lpFileName = Trim$(lpFileName)
   
   Else                                              ' do some adjustments on 'save file'
     
     For cn = 1 To Len(lpFileName)                   ' discard 0s
      If Mid$(lpFileName, cn, 1) = vbNullChar Then Mid$(lpFileName, cn, 1) = " "
     Next
     lpFileName = Trim$(lpFileName)
     
     Dim bCompleted As Boolean
     For cn = 1 To Len(lpFileName)                   ' check for extention
      If Mid$(lpFileName, cn, 1) = "." Then bCompleted = True
     Next
     If bCompleted = False Then                      ' add extention
      lpFileName = lpFileName & DefFileFilter
     End If
 
   End If
   
     '...
  ElseIf rval = 0 Then                               ' use cancel
     lpFileName = "none"
     '...
  Else                                               ' extended error
     'Call ErrorMsg("Error opening file!")
     '...
  End If
 End With
   
 strLastDir = lpFileName
   
End Sub


