Attribute VB_Name = "mDirectDraw"
Option Explicit
'-------------------------------------
'--> DirectDraw Engine
'--> by Peter "Pro-XeX" Petrov
'--> KenamicK Entertainment 1998-2002
'-------------------------------------

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Enum cnstSurfMemoryLoad
 SML_VIDEO = 0
 SML_SYSTEM
 SML_DEFAULT
End Enum

'Private Enum cnstLOADSOURCE
' LS_FROMFILE = 0
' LS_FROMRESOURCE
' LS_FROMBINRES
'End Enum

'' list of supportted DDraw Rops
'Public Type stDDRops
' DDAND    As Boolean
' DDOR     As Boolean
' DDINVERT As Boolean
'End Type

Private Const BMP_HEADER = &H4D42 ' BM = 19778
Private Const SRCCOPY = &HCC0020
Private Const DIB_RGB_COLORS = 0
Private Const GDI_ERROR = &HFFFF

'Bitmap file format structures
Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type

Public Type typeGFX_HBM                             ' I strongly recommend you
 dds As DirectDrawSurface7                          ' using a structure like this
 cx As Integer                                      ' for the Art. It gives u
 cy As Integer                                      ' a lot of flexability.
 bTrans As Boolean
 'hdc as Long
End Type

Public lpDX As DirectX7                         ' main DirectX object
Public lpDD          As DirectDraw7             ' DirectDraw object
Public lpPrim        As DirectDrawSurface7      ' primary surface
Public lpBack        As DirectDrawSurface7      ' backbuffer
Public sDDrawDriver  As String                  ' driver to use
Private lpGamma      As DirectDrawGammaControl  ' gamme object
Private lpGammaRamp0 As DDGAMMARAMP             ' original gamma ramp
Private lpGammaRamp1 As DDGAMMARAMP             ' custom gamma ramp
Private lpGammaRamp2 As DDGAMMARAMP             ' current gamma ramp
Private lpClipper As DirectDrawClipper          ' clipper for windowed mode
Private hw        As DDCAPS                     ' hardware capabilities
Private hel       As DDCAPS                     ' software capabilities

' Additional Declares
Private nBackBufferCount As Integer             ' backbuffers
Private bytBpp           As Byte
'Public lMemMethod As Long                      ' where to create surface (VID or SYS mem)
Public bColorFill        As Boolean             ' does device support color fills
Public bHardwareRasters  As Boolean             ' does device support hardware raster operations
Public bGamma            As Boolean             ' does device support gamma correction


' //////////////////////////////////////////////////////////
' //// Initialize Master DirectX Object
' //////////////////////////////////////////////////////////
Public Sub _
DXInit()                             ' initialize DirectX object
  
 On Local Error GoTo DXERROR:
  
  Set lpDX = New DirectX7
  If Err.Number <> 0 Then
     Call AppendToLog("Error: No DirectX version 7 present!")
     MakeError ("Game requires DirectX 7 or higher.")
  End If

Exit Sub

DXERROR:
 AppendToLog (DDGetErrorDesc(Err.Number))
 MakeError ("Game requires DirectX7 or greater.")
End Sub

' //////////////////////////////////////////////////////////
' //// Get Hardware Capabilites
' //////////////////////////////////////////////////////////
Public Sub _
DDGetCaps()
 
 AppendToLog ("Retrieving DirectDraw capabilites...")
 ' get capabilites
 Call lpDD.GetCaps(hw, hel)
  
 If (hw.lCaps And DDCAPS_NOHARDWARE) Then
  AppendToLog ("No hardware support at all.")
 End If
 
 ' check for color fills support
 If (hw.lCaps And DDCAPS_BLTCOLORFILL) Then
  AppendToLog ("Device supports color fills.")
  bColorFill = True
 Else
  AppendToLog ("Device does not support color fills.")
  bColorFill = False
 End If
 
 ' check for color keys
 If (hw.lCaps And DDCAPS_COLORKEY) Then
  If (hw.lCaps And DDCKEYCAPS_SRCBLT) Then
   AppendToLog ("Hardware src color key is possible.")
  End If
 Else
  AppendToLog ("Cannot do color keys in hardware!")
 End If
 
 If (hw.lCaps2 And DDCAPS2_PRIMARYGAMMA) Then
  bGamma = True
  AppendToLog ("Hardware gamma correction is possible.")
 Else
  bGamma = False
  AppendToLog ("Hardware gamma correction is not possible.")
 End If
 
 ' check for hardware raster operations
 'If (hw.lCaps And DDCAPS_COLORKEY) Then
 ' If ((hw.lRops And SRCCOPY) And (hw.lRops And SRCAND) And _
 '     (hw.lRops And SRCINVERT) And (hw.lRops And SRCPAINT)) Then
 '  bHardwareRasters = True
 '  AppendToLog ("Hardware rasters possible.")
 ' End If
 'Else
  bHardwareRasters = False
 ' AppendToLog ("Cannot do hardware rasters!")
 'End If
 
 ' stream video memory size
 AppendToLog ("Total Video Memory: " & (hw.lVidMemTotal / (1024&)) & " bytes")
 AppendToLog ("Free Video Memory: " & (hw.lVidMemFree / (1024&)) & " bytes")
 'If lMemMethod = DDSCAPS_VIDEOMEMORY Then
 ' AppendToLog ("Using VideoMemory for graphics")
 'ElseIf lMemMethod = DDSCAPS_SYSTEMMEMORY Then
 ' AppendToLog ("Using SystemMemory for graphics")
 'End If
 
 ' set backbuffers
 If (nBackBufferCount = 0 And (Not bWindowed)) Then
   If (hw.lVidMemTotal < (1024& * 1024&) * (bytBpp / 8)) Then
    AppendToLog ("Double Buffering. (not enough memory)")
    nBackBufferCount = 2
   Else
    AppendToLog ("Triple Buffering possible.")
    nBackBufferCount = 3
   End If
 End If

End Sub

' //////////////////////////////////////////////////////////
' //// Get video memory left
' //////////////////////////////////////////////////////////
Public Sub _
DDFreeMemToLog()
 
 Call lpDD.GetCaps(hw, hel)
 AppendToLog ("Free Video Memory: " & (hw.lVidMemFree / (1024&)) & " bytes")
End Sub

' //////////////////////////////////////////////////////////
' //// Init DDraw
' //////////////////////////////////////////////////////////
Public Sub _
DDInit(hwnd As Long, nWidth As Integer, nHeight As Integer, nBPP As Integer)
On Local Error GoTo DDERROR
  
  Call AppendToLog(LOG_DASH)
  Call AppendToLog("Opening DirectDraw...")
  If (sDDrawDriver = "") Then
   Call AppendToLog("Using Driver: Default")
  Else
   Call AppendToLog("Using Driver:" & sDDrawDriver)
  End If
  Set lpDD = lpDX.DirectDrawCreate(sDDrawDriver)          ' init Main Object

 'lMemMethod = DDSCAPS_SYSTEMMEMORY 'DDSCAPS_VIDEOMEMORY
  Call DDGetCaps
  bytBpp = CByte(nBPP)
  Call DDInitBuffers(hwnd, nWidth, nHeight)
  
Exit Sub

DDERROR: AppendToLog (DDGetErrorDesc(Err.Number))
         MakeError (DDGetErrorDesc(Err.Number))
End Sub
   
' //////////////////////////////////////////////////////////
' //// Create Primary and Back buffers
' //////////////////////////////////////////////////////////
Public Sub _
DDInitBuffers(hwnd As Long, nWidth As Integer, nHeight As Integer)

 On Local Error GoTo DDERROR

 Dim ddsd1 As DDSURFACEDESC2
 Dim ddsd2 As DDSURFACEDESC2
 Dim dcaps As DDSCAPS2
  
  If (Not bWindowed) Then
     
     AppendToLog ("Setting cooperativelevel...")
     Call lpDD.SetCooperativeLevel(hwnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWMODEX)
     AppendToLog ("Switching DisplayMode to " & nWidth & "x" & nHeight & "x" & bytBpp)
     Call lpDD.SetDisplayMode(nWidth, nHeight, CLng(bytBpp), 0, DDSDM_DEFAULT)
     DoEvents
     
     AppendToLog ("Creating Primary Surface...")
     With ddsd1                                 ' creating primary and backbuffers
        .ddsCaps.lCaps = DDSCAPS_FLIP Or DDSCAPS_COMPLEX Or DDSCAPS_PRIMARYSURFACE ' Or DDSCAPS_VIDEOMEMORY  ' or   'Or DDSCAPS_SYSTEMMEMORY
        .lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
        .lBackBufferCount = nBackBufferCount - 1
     End With
     Set lpPrim = lpDD.CreateSurface(ddsd1)
     AppendToLog ("Creating BackBuffer(s)...")
     dcaps.lCaps = DDSCAPS_BACKBUFFER
     Set lpBack = lpPrim.GetAttachedSurface(dcaps)
     lpBack.GetSurfaceDesc ddsd2
        
  Else
     AppendToLog ("Setting cooperativelevel...")
     Call lpDD.SetCooperativeLevel(hwnd, DDSCL_NORMAL) ' Or DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN)
     'Call lpDD.SetDisplayMode(MAX_CX, MAX_CY, BPP, 0, DDSDM_DEFAULT)
     AppendToLog ("Creating Primary Surface...")
     With ddsd1                                      ' create prim surface
       .lFlags = DDSD_CAPS
       .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
     End With
     Set lpPrim = lpDD.CreateSurface(ddsd1)
     
     AppendToLog ("Creating BackBuffer...")
     With ddsd2                                      ' create backbuffer
       .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
       .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN 'Or DDSCAPS_VIDEOMEMORY
       .lWidth = MAX_CX
       .lHeight = MAX_CY
     End With
     Set lpBack = lpDD.CreateSurface(ddsd2)
     
     Call AppendToLog("Setting Clipper...")
     Set lpClipper = lpDD.CreateClipper(0)            ' create and set clipper
     Call lpClipper.SetHWnd(hwnd)
     Call lpPrim.SetClipper(lpClipper)
  End If

  Call AppendToLog("DirectDraw was opened successfully.")
 
Exit Sub

DDERROR: AppendToLog (DDGetErrorDesc(Err.Number))
         MakeError (DDGetErrorDesc(Err.Number))
End Sub

' //////////////////////////////////////////////////////////
' //// Init Gamma object
' //////////////////////////////////////////////////////////
Public Sub _
DDInitGamma(Optional nR_old As Integer = MAX_INT, _
            Optional nG_old As Integer = MAX_INT, _
            Optional nB_old As Integer = MAX_INT)

 On Local Error GoTo DDERROR
  
 ' check for gamma support
 If (Not bGamma) Then Exit Sub
  
 ' create gamma
 AppendToLog ("Initializing gamma control...")
 Set lpGamma = lpPrim.GetDirectDrawGammaControl()
 ' get original gamma ramp
 AppendToLog ("Getting original ramp...")
 Call lpGamma.GetGammaRamp(DDSGR_DEFAULT, lpGammaRamp0)
 AppendToLog ("Gamma control initialized successfully.")
 ' reset custom gamma
 lpGammaRamp1 = lpGammaRamp0
   
Exit Sub

DDERROR: AppendToLog (DDGetErrorDesc(Err.Number))
         MakeError (DDGetErrorDesc(Err.Number))
End Sub

' //////////////////////////////////////////////////////////
' //// RestSet Gammaramp values
' //// nRed, nGreen, nBlue - colors intensity
' //////////////////////////////////////////////////////////
Public Sub _
DDReSetGamma()

 ' reset custom gamma
 lpGammaRamp1 = lpGammaRamp0
 
 ' assign new values
 Call lpGamma.SetGammaRamp(DDSGR_DEFAULT, lpGammaRamp1)

End Sub


' //////////////////////////////////////////////////////////
' //// Set Gamma values
' //// nRed, nGreen, nBlue - colors intensity
' //////////////////////////////////////////////////////////
Public Sub _
DDSetGamma(nRed As Integer, nGreen As Integer, nBlue As Integer)

 On Local Error GoTo DDERROR
  
 ' check for gamma support
 If (Not bGamma) Then Exit Sub

 Dim cn As Integer
 
 For cn = 0 To 255
  ' set red
  If (nRed < 0) Then lpGammaRamp2.red(cn) = _
  CSigned(CUnSigned(lpGammaRamp1.red(cn)) * (100 - Abs(nRed)) / 100)
  If (nRed = 0) Then lpGammaRamp2.red(cn) = lpGammaRamp1.red(cn)
  If (nRed > 0) Then lpGammaRamp2.red(cn) = _
  CSigned(65535 - ((65535 - CUnSigned(lpGammaRamp1.red(cn))) * (100 - nRed) / 100))
  ' set green
  If (nGreen < 0) Then lpGammaRamp2.green(cn) = _
  CSigned(CUnSigned(lpGammaRamp1.green(cn)) * (100 - Abs(nGreen)) / 100)
  If (nGreen = 0) Then lpGammaRamp2.green(cn) = lpGammaRamp1.green(cn)
  If (nGreen > 0) Then lpGammaRamp2.green(cn) = _
  CSigned(65535 - ((65535 - CUnSigned(lpGammaRamp1.green(cn))) * (100 - nGreen) / 100))
  ' set blue
  If (nBlue < 0) Then lpGammaRamp2.blue(cn) = _
  CSigned(CUnSigned(lpGammaRamp1.blue(cn)) * (100 - Abs(nBlue)) / 100)
  If (nBlue = 0) Then lpGammaRamp2.blue(cn) = lpGammaRamp1.blue(cn)
  If (nBlue > 0) Then lpGammaRamp2.blue(cn) = _
  CSigned(65535 - ((65535 - CUnSigned(lpGammaRamp1.blue(cn))) * (100 - nBlue) / 100))
 Next
 
 ' assign new values
 Call lpGamma.SetGammaRamp(DDSGR_DEFAULT, lpGammaRamp2)

Exit Sub

DDERROR: AppendToLog ("SETGAMMA:" & DDGetErrorDesc(Err.Number))
         MakeError (DDGetErrorDesc(Err.Number))
End Sub

' //////////////////////////////////////////////////////////
' //// Create Surface from Given Surface and dimensions
' //////////////////////////////////////////////////////////
Public Sub SurfaceFromSurface(lpSrcSurf As DirectDrawSurface7, _
                              lWidth As Integer, lHeight As Integer, _
                              lpDestSurf As DirectDrawSurface7, _
                              TRANSPARENT As Boolean, _
                              Optional TransIndex As Long = 0)
Dim ddsd1      As DDSURFACEDESC2
Dim ck         As DDCOLORKEY
Dim rDest      As RECT

   ' reset dest
   'Set lpDestSurf = Nothing
   
   With ddsd1
     .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
     .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
     .lWidth = lWidth
     .lHeight = lHeight
   End With
   
   ' create empty dest surface
   Set lpDestSurf = lpDD.CreateSurface(ddsd1)
   
   ' copy conatins from the source
   Call SetRect(rDest, 0, 0, lWidth, lHeight)
   'Call lpDestSurf.Blt(rDest, lpSrcSurf, rDest, DDBLT_WAIT)
   Call lpDestSurf.BltFast(0, 0, lpSrcSurf, rDest, DDBLTFAST_WAIT)
   
   If (TRANSPARENT) Then
      ck.low = TransIndex
      ck.high = TransIndex
      Call lpDestSurf.SetColorKey(DDCKEY_SRCBLT, ck)
   End If

End Sub

' //////////////////////////////////////////////////////////
' //// Creates an empty surface
' //// INT nWidth - width of the surface
' //// INT nHeight - height of the surface
' //////////////////////////////////////////////////////////
Public Function _
CreateEmptySurface(nWidth As Integer, nHeight As Integer) As DirectDrawSurface7

 On Local Error GoTo DDERROR                      ' create an empty surface

 Dim ddsd As DDSURFACEDESC2
    
    With ddsd
      .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
      .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
      .lWidth = nWidth
      .lHeight = nHeight
    End With
    ' create the surface
    Set CreateEmptySurface = lpDD.CreateSurface(ddsd)
     
Exit Function

DDERROR: AppendToLog (DDGetErrorDesc(Err.Number))
         MakeError (DDGetErrorDesc(Err.Number))
End Function

' //////////////////////////////////////////////////////////
' //// Create Surface from file
' //// STRING  lpF_Name - file name
' //// INT     nWidth - surface width ( 0 for default )
' //// INT     nHeight - surface height ( 0 for default )
' //// OPT.BOOL bTransparent - transparent surface ?
' //// OPT.LONG TransIndex - transparent color
' //// cnstMemMode - where to put it...
' //////////////////////////////////////////////////////////
Public Function _
DDLoadSurfaceFromFile(lpF_Name As String, _
                      nWidth As Integer, nHeight As Integer, _
                      bTransparent As Boolean, _
                      Optional TransIndex As Long = -1, _
                      Optional cnstMemMode As cnstSurfMemoryLoad = SML_DEFAULT) As DirectDrawSurface7
                         
 On Local Error GoTo DDERROR                                 ' trap loading errors
 
 Dim ddsd1      As DDSURFACEDESC2
 Dim ck         As DDCOLORKEY
 Dim TempSurf   As DirectDrawSurface7
 Dim lMemMethod As Long
 
 ' determine where to create the surface
 Select Case (cnstMemMode)
 
   ' --- load in system memory
   Case SML_SYSTEM
     lMemMethod = DDSCAPS_SYSTEMMEMORY
   ' --- load in video memory
   Case SML_VIDEO
     lMemMethod = DDSCAPS_VIDEOMEMORY
   ' --- defualt(s)
   Case SML_DEFAULT
     lMemMethod = 0
   Case Else
     lMemMethod = 0
 End Select
   
 ' clear temporary surface
 Set TempSurf = Nothing
 
 ' fill surface description
 With ddsd1
   .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or lMemMethod
   ' see if custom resizing should be done
   If (nWidth = 0 And nHeight = 0) Then
    .lFlags = DDSD_CAPS
   Else
   .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
   .lWidth = nWidth
   .lHeight = nHeight
   End If
 End With
 
 ' load surface
 Set TempSurf = lpDD.CreateSurfaceFromFile(lpF_Name, ddsd1)

 ' check for transprency
 If (bTransparent) Then
   
   If (TransIndex = -1) Then
    ' use color at position (1,1) for color key
    Call SetColorKeyAuto(TempSurf)
   
   ElseIf (TransIndex = 0) Then
    ' use black color key
    ck.low = TransIndex
    ck.high = TransIndex
    Call TempSurf.SetColorKey(DDCKEY_SRCBLT, ck)
   
   Else
    ' other colors need to be translated in app. for the color mode format
    Call SetColorKeyEx(TempSurf, TransIndex)
   End If
 End If

 ' return create surface
 Set DDLoadSurfaceFromFile = TempSurf
 ' destroy temporary surface
 Set TempSurf = Nothing

Exit Function

DDERROR: AppendToLog ("DD_ERROR: Could not load " & lpF_Name) ' whoopsy....
         MakeError ("Could not load " & lpF_Name)
End Function

' /////////////////////////////////////////////////////////////////
' //// Create surface(sprite) structure
' //// STRING  lpF_Name - bitmap file name/BinRes Container/Res ID
' //// INT     nWidth - surface width ( 0 for default )
' //// INT     nHeight - surface height ( 0 for default )
' //// OPT.BOOL bTransparent - transparent surface ?
' //// OPT.LONG TransIndex - transparent color
' //// cnstMemMode - where to put it...
' //////////////////////////////////////////////////////////
' /////////////////////////////////////////////////////////////////
Public Function _
CreateGFX_HBM(lpF_Name As String, _
              nWidth As Integer, nHeight As Integer, _
              bTransparent As Boolean, _
              Optional lTransIndex As Long = -1, _
              Optional cnstLR As cnstLOADSOURCE = LS_FROMFILE, _
              Optional lOffset As Long = 0, _
              Optional cnstMemMode As cnstSurfMemoryLoad = SML_DEFAULT _
              ) As typeGFX_HBM
                         
 ' trap erros
 'On Local Error GoTo DDERROR
  
 ' some temp values
 Dim TempSurf As DirectDrawSurface7
 Dim ddsd1 As DDSURFACEDESC2
 
 Select Case (cnstLR)
 
   ' --- load from file
   Case LS_FROMFILE
     Set TempSurf = DDLoadSurfaceFromFile(lpF_Name, nWidth, nHeight, _
                                          bTransparent, lTransIndex, _
                                          cnstMemMode)
   ' --- load from binary packet
   Case LS_FROMBINRES
     'If CKdfGfx.GetEntryPositionFromName(lpF_Name) = -1 Then Stop
     Set TempSurf = DDLoadBitmapFromBinRes(CKdfGfx.GetPacketName, CKdfGfx.GetEntryPositionFromName(lpF_Name), _
                                           nWidth, nHeight, _
                                           bTransparent, lTransIndex, _
                                           cnstMemMode, lpF_Name)
     'Set TempSurf = DDLoadBitmapFromBinRes(App.Path & "\gfx\pak\" & lpF_Name, 1, _
                                           nWidth, nHeight, _
                                           bTransparent, lTransIndex, _
                                           cnstMemMode, lpF_Name)
   
   
   ' --- load from resource
   Case LS_FROMRESOURCE
   '...
  
   ' error selection
   Case Else
     Debug.Print "Error GFX_LOAD case!"
     GoTo DDERROR
 
 End Select
  
 ' get surface description
 Call TempSurf.GetSurfaceDesc(ddsd1)
  
 ' fill info
 With CreateGFX_HBM
  Set .dds = Nothing
  Set .dds = TempSurf
  .cx = ddsd1.lWidth
  .cy = ddsd1.lHeight
  .bTrans = bTransparent
 End With
 
 ' destroy temp surface
 Set TempSurf = Nothing

Exit Function

DDERROR: 'AppendToLog ("DD_ERROR: Could not load " & lpF_Name) ' whoopsy....
         'MakeError ("Could not load " & lpF_Name)
End Function
                    
' //////////////////////////////////////////////////////////
' //// See how a color is represeted and apply it as a ck
' //////////////////////////////////////////////////////////
Public Sub _
SetColorKeyEx(lpSurf As DirectDrawSurface7, cr As Long) ' set surface color depending on it's background

 Dim ddsd As DDSURFACEDESC2
 Dim tClr As Long                                      ' tempcolor
 Dim ck As DDCOLORKEY
 Dim hTempDC As Long

 ' plot desired ck to the dc
 lpSurf.restore
 hTempDC = lpSurf.GetDC()
   SetPixel hTempDC, 1, 1, cr
 lpSurf.ReleaseDC hTempDC
 
 ' see how is the pixel represented in this color-mode
 Call lpSurf.GetSurfaceDesc(ddsd)                     ' get description
 lpSurf.Lock rEmpty, ddsd, DDLOCK_WAIT Or DDLOCK_READONLY Or DDLOCK_NOSYSLOCK, 0
 tClr = lpSurf.GetLockedPixel(1, 1)                   ' get color at position 0,0
 lpSurf.Unlock rEmpty                                 ' ...hey, where're you goin'? Unlock me... :)
 
 ck.high = tClr
 ck.low = tClr
 lpSurf.SetColorKey DDCKEY_SRCBLT, ck
 
End Sub

' //////////////////////////////////////////////////////////
' //// Set CK dependant on surfaces xy(1,1) coords
' //////////////////////////////////////////////////////////
Public Sub _
SetColorKeyAuto(lpSurf As DirectDrawSurface7)  ' set surface color depending on it's background
On Local Error GoTo DDErr
Dim ddsd As DDSURFACEDESC2
Dim tClr As Long                                      ' tempcolor
Dim ck As DDCOLORKEY
Dim hTempDC As Long

 Call lpSurf.GetSurfaceDesc(ddsd)                     ' get description
 lpSurf.Lock rEmpty, ddsd, DDLOCK_WAIT Or DDLOCK_READONLY Or DDLOCK_NOSYSLOCK, 0
 tClr = lpSurf.GetLockedPixel(1, 1)                   ' get color at position 0,0
 lpSurf.Unlock rEmpty                                 ' ...hey, where're you goin'? Unlock me... :)
 
 ck.high = tClr
 ck.low = tClr
 lpSurf.SetColorKey DDCKEY_SRCBLT, ck
Exit Sub

DDErr:
 lpSurf.Unlock rEmpty
End Sub
                        
' //////////////////////////////////////////////////////////
' //// Flip or Blit backbuffer onto primary surface
' //////////////////////////////////////////////////////////
Public Sub _
DDBlitToPrim()                         ' flip or blit backbuffer contents to primary
 Dim rSrc As RECT, rDest As RECT
 Dim ddrval As Long
 Dim lpPt As POINTAPI
  
 Call CheckIfTasked
 
 If (Not bWindowed) Then
    lpPrim.Flip Nothing, False ' DDFLIP_WAIT
    'If Err.Number = DDERR_WASSTILLDRAWING Then Debug.Print "DDERR_WASSTILDRAWING"
 Else
     Call ClientToScreen(frmMain.hwnd, lpPt)
     Call SetRect(rSrc, 0, 0, MAX_CX, MAX_CY)
     Call SetRect(rDest, lpPt.x, lpPt.y, lpPt.x + MAX_CX, lpPt.y + MAX_CY)
     'Call lpPrim.Blt(rEmpty, lpBack, rEmpty, DDBLT_WAIT)
     'Do While (1)
      ddrval = lpPrim.Blt(rDest, lpBack, rSrc, False)
     ' If (ddrval = DD_OK) Then Exit Do
     ' If (ddrval = DDERR_SURFACELOST) Then lpDD.RestoreAllSurfaces
     ' If (ddrval <> DDERR_WASSTILLDRAWING) Then Exit Do
     'Loop
 End If

End Sub

' //////////////////////////////////////////////////////////
' //// Restore display and cooperative modes
' //////////////////////////////////////////////////////////
Public Sub _
DDRestoreModes(hwnd As Long)
 
 Call lpDD.RestoreDisplayMode
 Call lpDD.SetCooperativeLevel(hwnd, DDSCL_NORMAL)
End Sub
 
' /////////////////////////////////////////////////////////////////
' //// Loads a bitmap from a given position in a binary data packet
' //// lpSurf - non-initialized DD7 surface
' //// lOffset - the position in the binary data packet
' //// INT     nWidth - surface width ( 0 for default )
' //// INT     nHeight - surface height ( 0 for default )
' //// OPT.BOOL bTransparent - transparent surface ?
' //// OPT.LONG TransIndex - transparent color
' //// cnstMemMode - where to put it...
' /////////////////////////////////////////////////////////////////
Public Function _
DDLoadBitmapFromBinRes(lpszLibrary As String, _
                       lOffset As Long, _
                       nWidth As Integer, nHeight As Integer, _
                       bTransparent As Boolean, _
                       Optional TransIndex As Long = -1, _
                       Optional cnstMemMode As cnstSurfMemoryLoad = SML_DEFAULT, _
                       Optional ss As String) As DirectDrawSurface7

  'Erase bmpInfo.bmiColors
  On Local Error GoTo DDERROR
  
  Dim bmpHeader As BITMAPFILEHEADER   ' Holds the file header
  Dim bmpInfo   As BITMAPINFO         ' Holds the bitmap info
  Dim bmpData() As Byte               ' Holds the pixel data
  
  Dim cn         As Integer           ' local counter
  Dim rval       As Long
  Dim nFF        As Integer
  Dim nlWidth     As Long
  Dim nlHeight    As Long
  
  ' get free file handle
  nFF = FreeFile()
  
  ' --- Open File ---
  Open (lpszLibrary) For Binary Access Read Lock Write As nFF
  ' get fileheader
  Get nFF, lOffset, bmpHeader
  
  ' check for bitamp header
  If (bmpHeader.bfType <> BMP_HEADER) Then
   'DDLoadBitmapFromBinRes = False
   GoTo DDERROR
   Exit Function
  End If
  
  ' get infoheader
  Get nFF, , bmpInfo.bmiHeader
  
  If (bmpInfo.bmiHeader.biClrUsed <> 0) Then
    For cn = 0 To bmpInfo.bmiHeader.biClrUsed - 1
     Get nFF, , bmpInfo.bmiColors(cn).rgbBlue
     Get nFF, , bmpInfo.bmiColors(cn).rgbGreen
     Get nFF, , bmpInfo.bmiColors(cn).rgbRed
     Get nFF, , bmpInfo.bmiColors(cn).rgbReserved
    Next cn
 
 ' --- setup 8 bit images ---
 ElseIf (bmpInfo.bmiHeader.biBitCount = 8) Then

    Get nFF, , bmpInfo.bmiColors
 End If

 ' size array
 If (bmpInfo.bmiHeader.biBitCount = 8) Then
    ReDim bmpData(FileSize(bmpInfo.bmiHeader.biWidth, _
                           bmpInfo.bmiHeader.biHeight))
    
 Else
  ReDim bmpData(bmpInfo.bmiHeader.biSizeImage - 1)
 End If
 
 ' get image
 Get nFF, , bmpData
  
 If (bmpInfo.bmiHeader.biBitCount = 8) Then
    bmpHeader.bfOffBits = 1078 ' 1024 + 54(header)
    With bmpInfo.bmiHeader
     .biSizeImage = FileSize(-bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight)
     .biClrUsed = 0
     .biClrImportant = 0
     .biXPelsPerMeter = 0
     .biYPelsPerMeter = 0
    End With
 End If
   
  ' --- Close File ---
  Close nFF
  
  
  ' now to stretch the (di)bits
  nlWidth = bmpInfo.bmiHeader.biWidth
  nlHeight = bmpInfo.bmiHeader.biHeight
  
  ' _-*-_*--- put to surface ---*_-*-_
  
  Dim lMemMethod As Long
  Dim TempSurf   As DirectDrawSurface7
  Dim ddsd1      As DDSURFACEDESC2
  Dim ck         As DDCOLORKEY
  
  ' determine where to create the surface
  Select Case (cnstMemMode)
 
    ' --- load in system memory
    Case SML_SYSTEM
      lMemMethod = DDSCAPS_SYSTEMMEMORY
    ' --- load in video memory
    Case SML_VIDEO
      lMemMethod = DDSCAPS_VIDEOMEMORY
    ' --- defualt(s)
    Case SML_DEFAULT
      lMemMethod = 0
    Case Else
      lMemMethod = 0
  End Select
   
  ' clear temporary surface
  Set TempSurf = Nothing
 
  ' fill surface description
  With ddsd1
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or lMemMethod
    ' see if defualt dimensions should be applied
    If (nWidth = 0 And nHeight = 0) Then
     .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
     .lWidth = nlWidth
     .lHeight = nlHeight
    Else
    .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    .lWidth = nWidth
    .lHeight = nHeight
    End If
  End With
 
  ' create surface
  Set TempSurf = lpDD.CreateSurface(ddsd1)
   
  ' copy to dc
  Dim hDestDC As Long
  
  hDestDC = TempSurf.GetDC()
    
  rval = StretchDIBits(hDestDC, 0, 0, ddsd1.lWidth, ddsd1.lHeight, _
                                0, 0, nlWidth, nlHeight, _
                                bmpData(0), _
                                bmpInfo, _
                                DIB_RGB_COLORS, _
                                SRCCOPY)
  ' done, now release
  Call TempSurf.ReleaseDC(hDestDC)
  
  ' check for success
  If (rval = GDI_ERROR) Then GoTo DDERROR

  ' check for transprency
  If (bTransparent) Then
   
   If (TransIndex = -1) Then
    ' use color at position (1,1) for color key
    Call SetColorKeyAuto(TempSurf)
    
   ElseIf (TransIndex = 0) Then
    ' use black color key
    ck.low = TransIndex
    ck.high = TransIndex
    Call TempSurf.SetColorKey(DDCKEY_SRCBLT, ck)
   
   Else
    ' other colors need to be translated in app. for the color mode format
    Call SetColorKeyEx(TempSurf, TransIndex)
   End If
  End If

  ' output surface
  Set DDLoadBitmapFromBinRes = TempSurf

  ' destroy temporary surface
  Set TempSurf = Nothing

Exit Function

DDERROR: AppendToLog ("DD_ERROR: Could not load from BinRes!")
         MakeError ("Error loading graphics from Data file!")
End Function

Private Function _
FileSize(lngWidth As Long, lngHeight As Long) As Long

    'Return the size of the image portion of the bitmap
    If lngWidth Mod 4 > 0 Then
        FileSize = ((lngWidth \ 4) + 1) * 4 * lngHeight - 1
    Else
        FileSize = lngWidth * lngHeight - 1
    End If

End Function
 
' //////////////////////////////////////////////////////////
' //// DirectDraw error descriptions
' //// lError - error number
' //////////////////////////////////////////////////////////
Public Function _
DDGetErrorDesc(lError As Long) As String

 Dim strMsg As String

 Select Case lError
     
   Case DDERR_ALREADYINITIALIZED
    strMsg = "The object has already been initialized."
   Case DDERR_BLTFASTCANTCLIP
    strMsg = "A DirectDrawClipper object is attached to a source surface that has passed into a call to the DirectDrawSurface7.BltFast method."
   Case DDERR_CANNOTATTACHSURFACE
    strMsg = "A surface cannot be attached to another requested surface."
   Case DDERR_CANNOTDETACHSURFACE
    strMsg = "A surface cannot be detached from another requested surface."
   Case DDERR_CANTCREATEDC
    strMsg = "Windows cannot create any more device contexts (DCs), or a DC was requested for a palette-indexed surface when the surface had no palette and the display mode was not palette-indexed (in this case DirectDraw cannot select a proper palette into the DC)."
   Case DDERR_CANTDUPLICATE
    strMsg = "Primary and 3-D surfaces, or surfaces that are implicitly created, cannot be duplicated."
   Case DDERR_CANTLOCKSURFACE
    strMsg = "Access to this surface is refused because an attempt was made to lock the primary surface without DCI support."
   Case DDERR_CANTPAGELOCK
    strMsg = "An attempt to page-lock a surface failed. Page lock does not work on a display-memory surface or an emulated primary surface."
   Case DDERR_CANTPAGEUNLOCK
    strMsg = "An attempt to page-unlock a surface failed. Page unlock does not work on a display-memory surface or an emulated primary surface."
   Case DDERR_CLIPPERISUSINGHWND
    strMsg = "An attempt was made to set a clip list for a DirectDrawClipper object that is already monitoring a window handle."
   Case DDERR_COLORKEYNOTSET
    strMsg = "No source color key is specified for this operation."
   Case DDERR_CURRENTLYNOTAVAIL
    strMsg = "No support is currently available."
   Case DDERR_DCALREADYCREATED
    strMsg = "A device context (DC) has already been returned for this surface. Only one DC can be retrieved for each surface."
   Case DDERR_DEVICEDOESNTOWNSURFACE
    strMsg = "Surfaces created by one DirectDraw device cannot be used directly by another DirectDraw device."
   Case DDERR_DIRECTDRAWALREADYCREATED
    strMsg = "A DirectDraw object representing this driver has already been created for this process."
   Case DDERR_EXCEPTION
    strMsg = "An exception was encountered while performing the requested operation."
   Case DDERR_EXCLUSIVEMODEALREADYSET
    strMsg = "An attempt was made to set the cooperative level when it was already set to exclusive."
   Case DDERR_EXPIRED
    strMsg = "The data has expired and is therefore no longer valid."
   Case DDERR_GENERIC
    strMsg = "There is an undefined error condition."
   Case DDERR_HEIGHTALIGN
    strMsg = "The height of the provided rectangle is not a multiple of the required alignment."
   Case DDERR_HWNDALREADYSET
    strMsg = "The DirectDraw cooperative level window handle has already been set. It cannot be reset while the process has surfaces or palettes created."
   Case DDERR_HWNDSUBCLASSED
    strMsg = " DirectDraw is prevented from restoring state because the DirectDraw cooperative level window handle has been subclassed."
   Case DDERR_IMPLICITLYCREATED
    strMsg = "The surface cannot be restored because it is an implicitly created surface."
   Case DDERR_INCOMPATIBLEPRIMARY
    strMsg = "The primary surface creation request does not match with the existing primary surface."
   Case DDERR_INVALIDCAPS
    strMsg = "One or more of the capability bits passed to the callback function are incorrect."
   Case DDERR_INVALIDCLIPLIST
    strMsg = "DirectDraw does not support the provided clip list."
   Case DDERR_INVALIDDIRECTDRAWGUID
    strMsg = "The globally unique identifier (GUID) passed to the DirectX7.DirectDrawCreate function is not a valid DirectDraw driver identifier."
   Case DDERR_INVALIDMODE
    strMsg = "DirectDraw does not support the requested mode."
   Case DDERR_INVALIDOBJECT
    strMsg = "DirectDraw received a pointer that was an invalid DirectDraw object."
   Case DDERR_INVALIDPARAMS
    strMsg = "One or more of the parameters passed to the method are incorrect."
   Case DDERR_INVALIDPIXELFORMAT
    strMsg = "The pixel format was invalid as specified."
   Case DDERR_INVALIDPOSITION
    strMsg = "The position of the overlay on the destination is no longer legal."
   Case DDERR_INVALIDRECT
    strMsg = "The provided rectangle was invalid."
   Case DDERR_INVALIDSTREAM
    strMsg = "The specified stream contains invalid data."
   Case DDERR_INVALIDSURFACETYPE
    strMsg = "The requested operation could not be performed because the surface was of the wrong type."
   Case DDERR_LOCKEDSURFACES
    strMsg = "One or more surfaces are locked."
   Case DDERR_MOREDATA
    strMsg = "There is more data available than the specified buffer size can hold."
   Case DDERR_NO3D
    strMsg = "No 3-D hardware or emulation is present."
   Case DDERR_NOALPHAHW
    strMsg = "No alpha acceleration hardware is present or available."
   Case DDERR_NOBLTHW
    strMsg = "No blitter hardware is present."
   Case DDERR_NOCLIPLIST
    strMsg = "No clip list is available."
   Case DDERR_NOCLIPPERATTACHED
    strMsg = "No DirectDrawClipper object is attached to the surface object."
   Case DDERR_NOCOLORCONVHW
    strMsg = "No color-conversion hardware is present or available."
   Case DDERR_NOCOLORKEY
    strMsg = "The surface does not currently have a color key."
   Case DDERR_NOCOLORKEYHW
    strMsg = "There is no hardware support for the destination color key."
   Case DDERR_NOCOOPERATIVELEVELSET
    strMsg = "A create function was called when the DirectDraw7.SetCooperativeLevel method had not been called."
   Case DDERR_NODC
    strMsg = "No DC has ever been created for this surface."
   Case DDERR_NODDROPSHW
    strMsg = " No DirectDraw raster operation (ROP) hardware is available."
   Case DDERR_NODIRECTDRAWHW
    strMsg = "Hardware-only DirectDraw object creation is not possible; the driver does not support any hardware."
   Case DDERR_NODIRECTDRAWSUPPORT
    strMsg = "DirectDraw support is not possible with the current display driver."
   Case DDERR_NOEMULATION
    strMsg = "Software emulation is not available."
   Case DDERR_NOEXCLUSIVEMODE
    strMsg = "The operation requires the application to have exclusive mode, but the application does not have exclusive mode."
   Case DDERR_NOFLIPHW
    strMsg = "Flipping visible surfaces is not supported."
   Case DDERR_NOFOCUSWINDOW
    strMsg = "An attempt was made to create or set a device window without first setting the focus window."
   Case DDERR_NOGDI
    strMsg = "No GDI is present."
   Case DDERR_NOHWND
    strMsg = "Clipper notification requires a window handle, or no window handle was previously set as the cooperative level window handle."
   Case DDERR_NOMIPMAPHW
    strMsg = "No mipmap-capable texture mapping hardware is present or available."
   Case DDERR_NOMIRRORHW
    strMsg = "No mirroring hardware is present or available."
   Case DDERR_NONONLOCALVIDMEM
    strMsg = "An attempt was made to allocate nonlocal video memory from a device that does not support nonlocal video memory."
   Case DDERR_NOOPTIMIZEHW
    strMsg = "The device does not support optimized surfaces."
   Case DDERR_NOOVERLAYHW
    strMsg = "No overlay hardware is present or available."
   Case DDERR_NOPALETTEATTACHED
    strMsg = "No palette object is attached to this surface."
   Case DDERR_NOPALETTEHW
    strMsg = "There is no hardware support for 16- or 256-color palettes."
   Case DDERR_NORASTEROPHW
    strMsg = " No appropriate raster operation hardware is present or available."
   Case DDERR_NOROTATIONHW
    strMsg = "No rotation hardware is present or available."
   Case DDERR_NOSTEREOHARDWARE
    strMsg = "No stereo hardware is present or available."
   Case DDERR_NOSTRETCHHW
    strMsg = "There is no hardware support for stretching."
   Case DDERR_NOSURFACELEFT
    strMsg = "No hardware is present that supports stereo surfaces."
   Case DDERR_NOT4BITCOLOR
    strMsg = "The DirectDrawSurface object is not using a 4-bit color palette, and the requested operation requires a 4-bit color palette."
   Case DDERR_NOT4BITCOLORINDEX
    strMsg = "The DirectDrawSurface object is not using a 4-bit color index palette, and the requested operation requires a 4-bit color index palette."
   Case DDERR_NOT8BITCOLOR
    strMsg = "The DirectDrawSurface object is not using an 8-bit color palette, and the requested operation requires an 8-bit color palette."
   Case DDERR_NOTAOVERLAYSURFACE
    strMsg = "An overlay component was called for a non-overlay surface."
   Case DDERR_NOTEXTUREHW
    strMsg = "No texture-mapping hardware is present or available."
   Case DDERR_NOTFLIPPABLE
    strMsg = "An attempt was made to flip a surface that cannot be flipped."
   Case DDERR_NOTFOUND
    strMsg = "The requested item was not found."
   Case DDERR_NOTINITIALIZED
    strMsg = "An attempt was made to call an interface method of a DirectDraw object created by CoCreateInstance before the object was initialized."
   Case DDERR_NOTLOADED
    strMsg = "The surface is an optimized surface, but it has not yet been allocated any memory."
   Case DDERR_NOTLOCKED
    strMsg = "An attempt was made to unlock a surface that was not locked."
   Case DDERR_NOTPAGELOCKED
    strMsg = "An attempt was made to page-unlock a surface with no outstanding page locks."
   Case DDERR_NOTPALETTIZED
    strMsg = "The surface being used is not a palette-based surface."
   Case DDERR_NOVSYNCHW
    strMsg = "There is no hardware support for vertical blank synchronized operations."
   Case DDERR_NOZBUFFERHW
    strMsg = "There is no hardware support for z-buffers."
   Case DDERR_NOZOVERLAYHW
    strMsg = "The hardware does not support z-ordering of overlays."
   Case DDERR_OUTOFCAPS
    strMsg = " The hardware needed for the requested operation has already been allocated."
   Case DDERR_OUTOFMEMORY
    strMsg = "DirectDraw does not have enough memory to perform the operation."
   Case DDERR_OUTOFVIDEOMEMORY
    strMsg = "DirectDraw does not have enough display memory to perform the operation."
   Case DDERR_OVERLAPPINGRECTS
    strMsg = "The source and destination rectangles are on the same surface and overlap each other."
   Case DDERR_OVERLAYCANTCLIP
    strMsg = "The hardware does not support clipped overlays."
   Case DDERR_OVERLAYCOLORKEYONLYONEACTIVE
    strMsg = "An attempt was made to have more than one color key active on an overlay."
   Case DDERR_OVERLAYNOTVISIBLE
    strMsg = "The method was called on a hidden overlay."
   Case DDERR_PALETTEBUSY
    strMsg = "Access to this palette is refused because the palette is locked by another thread."
   Case DDERR_PRIMARYSURFACEALREADYEXISTS
    strMsg = "This process has already created a primary surface."
   Case DDERR_REGIONTOOSMALL
    strMsg = "The region passed to the DirectDrawClipper.GetClipList method is too small."
   Case DDERR_SURFACEALREADYATTACHED
    strMsg = "An attempt was made to attach a surface to another surface to which it is already attached."
   Case DDERR_SURFACEALREADYDEPENDENT
    strMsg = "An attempt was made to make a surface a dependency of another surface on which it is already dependent."
   Case DDERR_SURFACEBUSY
    strMsg = "Access to the surface is refused because the surface is locked by another thread."
   Case DDERR_SURFACEISOBSCURED
    strMsg = "Access to the surface is refused because the surface is obscured."
   Case DDERR_SURFACELOST
    strMsg = "Access to the surface is refused because the surface memory is gone. Call the DirectDrawSurface7.Restore method on this surface to restore the memory associated with it."
   Case DDERR_SURFACENOTATTACHED
    strMsg = "The requested surface is not attached."
   Case DDERR_TOOBIGHEIGHT
    strMsg = "The height requested by DirectDraw is too large."
   Case DDERR_TOOBIGSIZE
    strMsg = "The size requested by DirectDraw is too large. However, the individual height and width are valid sizes."
   Case DDERR_TOOBIGWIDTH
    strMsg = "The width requested by DirectDraw is too large."
   Case DDERR_UNSUPPORTED
    strMsg = "The operation is not supported."
   Case DDERR_UNSUPPORTEDFORMAT
    strMsg = "The FourCC format requested is not supported by DirectDraw."
   Case DDERR_UNSUPPORTEDMASK
    strMsg = "The bitmask in the pixel format requested is not supported by DirectDraw."
   Case DDERR_UNSUPPORTEDMODE
    strMsg = "The display is currently in an unsupported mode."
   Case DDERR_VERTICALBLANKINPROGRESS
    strMsg = "A vertical blank is in progress."
   Case DDERR_VIDEONOTACTIVE
    strMsg = "The video port is not active."
   Case DDERR_WASSTILLDRAWING
    strMsg = "The previous blit operation that is transferring information to or from this surface is incomplete."
   Case DDERR_WRONGMODE
    strMsg = "This surface cannot be restored because it was created in a different mode."
   Case DDERR_XALIGN
    strMsg = "The provided rectangle was not horizontally aligned on a required boundary."
   'Case E_INVALIDINTERFACE
   ' strMsg = "The specified interface is invalid or does not exist."
   'Case E_OUTOFMEMORY
   ' strMsg = "Not enough free memory to complete the method."
   
   Case Else
    strMsg = "Error description not found!"
    
 End Select

 DDGetErrorDesc = "DDERROR:" & strMsg
 AppendToLog (strMsg)

End Function
 
' //////////////////////////////////////////////////////////
' //// Release DirectDraw objects
' //////////////////////////////////////////////////////////
Public Sub _
DDRelease()
 
 'AppendToLog (LOG_DASH)
 AppendToLog ("Closing DirectDraw")
 Set lpClipper = Nothing
 Set lpBack = Nothing
 Set lpPrim = Nothing
 Set lpDD = Nothing

End Sub

' //////////////////////////////////////////////////////////
' //// Release main DirectX object
' //////////////////////////////////////////////////////////
Public Sub _
DXRelease()
 Set lpDX = Nothing
End Sub

