Attribute VB_Name = "mGFX"
Option Explicit

'-------------------------------------
'--> GFX&Blitting Module
'-------------------------------------

Public Enum enumRasterOps                             ' API Raster Operations
  R_AND = SRCAND
  R_OR = SRCPAINT                                     ' for invisiblity
  R_XOR = SRCINVERT                                   ' nice clr effect
  R_COPY = SRCCOPY
End Enum

Public Enum enumBltMirror
  BM_LEFT = 0
  BM_RIGHT
End Enum

Private Type stRGB                   ' rgb sturcture
  r As Integer
  G As Integer
  B As Integer
End Type

Public m_gfxGammaRGB As stRGB        ' gamma colors
Private fnt          As New StdFont  ' local font info


' //////////////////////////////////////////////////////////
' //// Clear backbuffer contents
' //// BYTE bytSpeed - fading speed
' //////////////////////////////////////////////////////////
Public Sub _
GFXFadeInOut(Optional bytSpeed As Byte = 0)
 
 ' -99/+99
  
 Static bFade   As Boolean
 Static bFadeIn As Boolean
 Static nR As Integer, nG As Integer, nB As Integer, nspeed As Integer
 
 ' speed is set -> new fade call
 If (bytSpeed > 0) Then
  bFade = True
  bFadeIn = True
  nR = 0
  nG = 0
  nB = 0
  nspeed = CInt(bytSpeed)
 End If
 ' exit if no there's no fade call to update
 If (Not bFade) Then Exit Sub
  
 If (bFadeIn) Then
  
  nR = nR + nspeed
  nG = nG + nspeed
  nB = nG + nspeed
  Call mDirectDraw.DDSetGamma(nR, nG, nB)
  
  ' check vals
  If (nR > 90) Then bFadeIn = False
  
 ElseIf (Not bFadeIn) Then
 
  nR = nR - nspeed
  nG = nG - nspeed
  nB = nB - nspeed
  Call mDirectDraw.DDSetGamma(nR, nG, nB)
 
  ' check vals
  If (nR < (nspeed + nspeed)) Then
   bFade = False
   ' reset to custom ramp
   Call mDirectDraw.DDSetGamma(m_gfxGammaRGB.r, _
                               m_gfxGammaRGB.G, _
                               m_gfxGammaRGB.B)
  End If
  
 End If
 

End Sub


' //////////////////////////////////////////////////////////
' //// Clear backbuffer contents
' //////////////////////////////////////////////////////////
Public Sub _
GFXClearBackBuffer()

 Dim ddrval As Long
 ddrval = lpBack.BltColorFill(rEmpty, 0)
 
 ' try clearing using BltFX function
 If (ddrval <> DD_OK) Then
  Dim DDBLTFX As DDBLTFX
  DDBLTFX.lFill = 0
  
  ddrval = lpBack.BltFX(rEmpty, Nothing, rEmpty, DDBLT_COLORFILL, DDBLTFX)
  If (ddrval <> DD_OK) Then
   '...blit empty surface
  End If
 End If

End Sub

' //////////////////////////////////////////////////////////
' //// Draw text on a passed surface
' //////////////////////////////////////////////////////////
Public Sub _
GFXTextOut(lpSurf As DirectDrawSurface7, _
          ByVal x As Integer, ByVal y As Integer, _
          lpText As String, bytFontSize As Byte, _
          lColor As Long, _
          Optional lBackColor As Long = -1, _
          Optional bBold As Boolean = False)
                   
  
  On Local Error GoTo GFXTO_ERROR:
  
  If (Len(lpText) < 1) Then Exit Sub
  
  'Dim lhDC As Long
    
  fnt.name = "verdana"
  fnt.Size = bytFontSize \ 2
  fnt.Bold = bBold
  lpSurf.SetFont fnt
  
  If (lBackColor <> -1) Then
   lpSurf.SetFontTransparency False
   lpSurf.SetFontBackColor lBackColor
   lpSurf.SetForeColor lColor
   lpSurf.DrawText x, y, lpText, False
  Else
   lpSurf.SetFontTransparency True
   lpSurf.SetForeColor lColor
   lpSurf.DrawText x, y, lpText, False
  End If
    
  'lpSurf.restore
  'lhDC = lpSurf.GetDC()
  ' Call DrawText(lhDC, x, y, lpText, bytFontSize, bBold, lColor, lBackColor)
  'lpSurf.ReleaseDC lhDC
Exit Sub

GFXTO_ERROR:
 Debug.Print "Error getting text out!"
End Sub


' //////////////////////////////////////////////////////////
' //// Clip rect with another rect
' //////////////////////////////////////////////////////////
Public Sub _
ClipRect(r1 As RECT, r2 As RECT, r3 As RECT)
  r1.Top = r2.Top - r3.Top
  r1.Left = r2.Left - r3.Left
  r1.Right = r2.Right - r3.Left
  r1.Bottom = r2.Bottom - r3.Top
End Sub


' //////////////////////////////////////////////////////////
' //// Bltfast proc. that checks if a surface is in the
' //// world range and automatically cuts it if needed
' //////////////////////////////////////////////////////////
Public Sub _
BltFast(ByVal x, ByVal y, lpSurf As DirectDrawSurface7, rSrc As RECT, Trans As Boolean)
Dim lFlags As Long                                   ' Fast Blit Procedure
Dim rResult As RECT, rDest As RECT
'Dim off As Integer
Dim ddrval As Long

  If Trans = True Then                               ' see if transparent and apply flags
     lFlags = DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
     'lFlags = DDBLTFAST_SRCCOLORKEY
  Else
     lFlags = DDBLTFAST_WAIT
  End If
                                                     ' setup destination rectangle
  Call SetRect(rDest, x, y, x + rSrc.Right, y + rSrc.Bottom)
  
  If IntersectRect(rResult, rScreen, rDest) Then     ' see if surface overlaps current world_screen position
     Call ClipRect(rSrc, rResult, rDest)
     ddrval = lpBack.BltFast(rResult.Left, rResult.Top, lpSurf, rSrc, lFlags)
     'If ddrval = DDERR_WASSTILLDRAWING Then Debug.Print "WASSTILLDRAING_BLTFAST"
  End If
  
End Sub

' //////////////////////////////////////////////////////////
' //// Same as BltFast but a little more optimezed for .Top
' //// and .Left Rect structure members
' //////////////////////////////////////////////////////////
Public Sub _
BltFastW(ByVal x, ByVal y, lpSurf As DirectDrawSurface7, rSrc As RECT, Trans As Boolean)
Dim lFlags As Long                                   ' Fast Blit Procedure
Dim rResult As RECT, rDest As RECT
Dim xOff As Integer, yOff As Integer

  If Trans = True Then                               ' see if transparent and apply flags
     lFlags = DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
     'lFlags = DDBLTFAST_SRCCOLORKEY
  Else
     lFlags = DDBLTFAST_WAIT
  End If

  xOff = rSrc.Left                                   ' remember top and left offsets
  yOff = rSrc.Top
                                                     ' setup destination rectangle  Call SetRect(rDest, x, y, x + rSrc.Right, y + rSrc.Bottom)
  Call SetRect(rDest, x, y, x + rSrc.Right, y + rSrc.Bottom)
  If IntersectRect(rResult, rScreen, rDest) Then     ' see if surface overlaps current world_screen position
     Call ClipRect(rSrc, rResult, rDest)
     rSrc.Left = xOff
     rSrc.Top = yOff
     Call lpBack.BltFast(rResult.Left, rResult.Top, lpSurf, rSrc, lFlags)
  End If
  
End Sub


' //////////////////////////////////////////////////////////
' //// Blit a surface using a Raster Operation
' //////////////////////////////////////////////////////////
Public Function _
BltFX(ByVal x, ByVal y, lpSurf As DirectDrawSurface7, rSrc As RECT, RasterOp As enumRasterOps, Trans As Boolean) As Boolean

 Dim lFlags  As Long                                  ' FXBlit procedure
 Dim rResult As RECT, rDest As RECT
 Dim ddfx    As DDBLTFX
 Dim hDestDC As Long                                 ' Dest DC for GDI emulation
 Dim hSrcDC  As Long                                 ' source DC for GDI emulation
 Dim rval    As Long
  
  If Trans = True Then                               ' see if transparent and apply flags
     lFlags = DDBLT_WAIT Or DDBLT_KEYSRC
  Else
     lFlags = DDBLT_WAIT
  End If
   
  Call SetRect(rDest, x, y, x + rSrc.Right, y + rSrc.Bottom)  ' setup destination rectangle
  ddfx.lROP = RasterOp                               ' set desired raster opreand
  
  BltFX = True
  
  If IntersectRect(rResult, rScreen, rDest) Then     ' see if surface overlaps current world_screen position
     Call ClipRect(rSrc, rResult, rDest)
     ' try rops trough DirectDraw
     If (lpBack.BltFX(rResult, lpSurf, rSrc, DDBLT_ROP Or lFlags, ddfx)) <> DD_OK Then
      ' if failed then use GDI emulation
       ' get surfaces DC
       hDestDC = lpBack.GetDC()
       hSrcDC = lpSurf.GetDC()
       ' do blitting
       rval = BitBlt(hDestDC, rResult.Left, rResult.Top, rResult.Right, rResult.Bottom, _
                    hSrcDC, 0, 0, RasterOp)
       ' release surfaces
       Call lpSurf.ReleaseDC(hSrcDC)
       Call lpBack.ReleaseDC(hDestDC)
      ' check for blitting failure
      If (rval <> 0) Then BltFX = False
     End If
  End If
     
End Function


' //////////////////////////////////////////////////////////
' //// Blits a DD surface to another DD surface using a
' //// Raster Operation via GDI
' //////////////////////////////////////////////////////////
Public Function _
BltFXHel(ByVal x, ByVal y, lpSrcSurf As DirectDrawSurface7, _
         rSrc As RECT, RasterOp As enumRasterOps) As Boolean  ' ', Trans As Boolean)
 'Dim lFlags As Long
 Dim rResult As RECT, rDest As RECT
 Dim hdcDest As Long, hdcSrc As Long
 Dim rval As Long
 
  'If Trans = True Then                              ' see if transparent and apply flags
  '   lFlags = DDBLT_WAIT Or DDBLT_KEYSRC
  'Else
  '   lFlags = DDBLT_WAIT
  'End If
  
  Call SetRect(rDest, x, y, x + rSrc.Right, y + rSrc.Bottom)  ' setup destination rectangle
  
  If IntersectRect(rResult, rScreen, rDest) Then     ' see if surface overlaps current world_screen position
     Call ClipRect(rSrc, rResult, rDest)
     ' do blitting
     lpBack.restore
     hdcDest = lpBack.GetDC()
     hdcSrc = lpSrcSurf.GetDC()
     
      rval = BitBlt(hdcDest, rResult.Left, rResult.Top, _
                    rResult.Right, rResult.Bottom, _
                    hdcSrc, 0, 0, RasterOp)
     lpSrcSurf.ReleaseDC hdcSrc
     lpBack.ReleaseDC hdcDest
     ' see if blitting was successful
     If rval <> 0 Then BltFXHel = True Else BltFXHel = False
  End If
  
End Function


' //////////////////////////////////////////////////////////
' //// Blits the mirror image of a surface
' //////////////////////////////////////////////////////////
Public Sub _
BltMirror(ByVal x, ByVal y, lpSurf As DirectDrawSurface7, rSrc As RECT, Trans As Boolean)
Dim lFlags As Long                                   ' Mirror FXBlit procedure
Dim rResult As RECT, rDest As RECT, rTemp As RECT
Dim cx As Integer, cy As Integer
Dim ddfx As DDBLTFX

  If Trans = True Then                               ' see if transparent and apply flags
     lFlags = DDBLT_WAIT Or DDBLT_KEYSRC Or DDBLT_DDFX
  Else
     lFlags = DDBLT_WAIT
  End If
   
  Call SetRect(rDest, x, y, x + rSrc.Right, y + rSrc.Bottom)  ' setup destination rectangle
  cx = rSrc.Right                                    ' save surface dimensions
  cy = rSrc.Bottom
  ddfx.lDDFX = DDBLTFX_MIRRORLEFTRIGHT
  
  If IntersectRect(rResult, rScreen, rDest) Then     ' see if surface overlaps current world_screen position
     Call ClipRect(rSrc, rResult, rDest)
     Call CopyRect(rTemp, rSrc)                      ' copy to temp rect
     ' we need to reverse the source rectangle so the mirror surface blit propertly
     Call SetRect(rSrc, cx - rTemp.Right, _
                        cy - rTemp.Bottom, _
                        cx - rTemp.Left, _
                        cy - rTemp.Top)
     
     Call lpBack.BltFX(rResult, lpSurf, rSrc, lFlags, ddfx)
  End If
 
End Sub


' //////////////////////////////////////////////////////////
' //// Blits surface(sprite) structure
' //////////////////////////////////////////////////////////
Public Sub _
BltFastGFX_HBM(ByVal x, ByVal y, gfxHbm As typeGFX_HBM)
Dim lFlags As Long                                   ' GFX_HBM struct - FBP
Dim rResult As RECT, rDest As RECT, rSrc As RECT
Dim off As Integer
Dim ddrval As Long

 With gfxHbm
  If .bTrans = True Then                             ' see if transparent
     lFlags = DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
  Else
     lFlags = DDBLTFAST_WAIT
  End If
  
  Call SetRect(rSrc, 0, 0, .cx, .cy)
  Call SetRect(rDest, x, y, x + .cx, y + .cy)
  
  If IntersectRect(rResult, rScreen, rDest) Then     ' see if surface overlaps current world_screen position
     Call ClipRect(rSrc, rResult, rDest)
     ddrval = lpBack.BltFast(rResult.Left, rResult.Top, .dds, rSrc, lFlags)
     If ddrval = DDERR_WASSTILLDRAWING Then Debug.Print "WASSTILLDRAING_BLTFAST"
  End If
 End With
 
End Sub


' //////////////////////////////////////////////////////////
' //// Blits surface(sprite) structure with a RasterOp
' //////////////////////////////////////////////////////////
Public Sub _
BltFxGFX_HBM(ByVal x, ByVal y, gfxHbm As typeGFX_HBM, RasterOp As enumRasterOps)

Dim rSrc As RECT
  
 Call SetRect(rSrc, 0, 0, gfxHbm.cx, gfxHbm.cy)
 Call BltFX(x, y, gfxHbm.dds, rSrc, RasterOp, gfxHbm.bTrans)
 
End Sub


' //////////////////////////////////////////////////////////
' //// This is a little exaggerated.Actually we take a
' //// surface and set all it's pixels to a desired color.
' //// Later, you may blit it using a raster operation
' //// (SRC_AND) to achieve kind of a shadow effect.
' //////////////////////////////////////////////////////////
Public Sub _
CreateShadow(lpShadowSurf As DirectDrawSurface7)
On Local Error GoTo GFXError

 Dim rShadow As RECT
 Dim i As Integer, j As Integer                          ' local counters
 Dim ddsd_Shadow As DDSURFACEDESC2
 Dim Pict() As Byte                                      ' array that will hold raw data
 
 lpShadowSurf.GetSurfaceDesc ddsd_Shadow
 lpShadowSurf.Lock rShadow, ddsd_Shadow, DDLOCK_WAIT Or DDLOCK_WRITEONLY, 0
 lpShadowSurf.GetLockedArray Pict()
  
   For j = 0 To ddsd_Shadow.lHeight - 1
    For i = 0 To ddsd_Shadow.lWidth * 2 - 1 Step 2       ' we assume that displaymode is 16 bit
                                                         ' that means surface_width =(16/8)*surface_width
        If Pict(i, j) <> 0 Then Pict(i, j) = 8 'SHADOW_COLOR
        If Pict(i + 1, j) <> 0 Then Pict(i + 1, j) = 8 ' SHADOW_COLOR
    Next
   Next
 
 lpShadowSurf.Unlock rShadow

Exit Sub
GFXError:  lpShadowSurf.Unlock rShadow                   ' make sure we unlock
           AppendToLog ("Error creating Shadow...")
           MakeError ("Error createing Shadow!")
End Sub


' ------------------- Some GDI functions -----------------------

' //////////////////////////////////////////////////////////
' //// Blits a hDC to a DD surface using a Raster Operation
' //////////////////////////////////////////////////////////
Public Function _
BltFXGDI(ByVal x, ByVal y, lpSrcDC As Long, rSrc As RECT, RasterOp As enumRasterOps) As Boolean ' ', Trans As Boolean)
 Dim lFlags As Long                                  ' FXBlit procedure
 Dim rResult As RECT, rDest As RECT
 Dim hTempDC As Long
 Dim rval As Long
 
  'If Trans = True Then                              ' see if transparent and apply flags
     'lFlags = DDBLT_WAIT Or DDBLT_KEYSRC
  'Else
     'lFlags = DDBLT_WAIT
  'End If
   
  Call SetRect(rDest, x, y, x + rSrc.Right, y + rSrc.Bottom)  ' setup destination rectangle
  
  If IntersectRect(rResult, rScreen, rDest) Then     ' see if surface overlaps current world_screen position
     Call ClipRect(rSrc, rResult, rDest)
     ' do blitting
     lpBack.restore
     hTempDC = lpBack.GetDC()
      rval = BitBlt(hTempDC, rResult.Left, rResult.Top, rResult.Right, rResult.Bottom, _
                    lpSrcDC, 0, 0, RasterOp)
     lpBack.ReleaseDC hTempDC
     ' see if blitting was successful
     If rval <> 0 Then BltFXGDI = True Else BltFXGDI = False
  End If
  
End Function


' //////////////////////////////////////////////////////////
' //// This will load an image into a hDC
' //////////////////////////////////////////////////////////
Public Function _
CreateImageDC(SrcDC As Long, lpszFileName As String, nWidth As Integer, nHeight As Integer) As Long
 Dim hbm As Long
 Dim hInst As Long
 Dim oldObject As Long, NewObject As Long
  
 SrcDC = CreateCompatibleDC(ByVal 0&)
 ' check for failure
 If SrcDC = 0 Then
  AppendToLog ("GDI Error: Error creating DC !")
  MakeError ("Error loading " & lpszFileName & " !")
  Exit Function
 End If
 
 hInst = App.hInstance
 ' load the image
 hbm = LoadImage(hInst, lpszFileName, _
                 IMAGE_BITMAP, nWidth, nHeight, _
                 LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
 ' check for failure
 If hbm = 0 Then
  Call DeleteDC(SrcDC)
  AppendToLog ("GDI Error: Error loading  bitmap image!")
  MakeError ("Error loading " & lpszFileName & " !")
  Exit Function
 End If
 oldObject = SelectObject(SrcDC, hbm)
 
 ' delete bitmap
 If DeleteObject(hbm) = 0 Then
  AppendToLog ("GDI Error: Error deleting bitmap image!")
  'MakeError ("Error loading " & lpszFileName & " !")
 End If
  
End Function


'///////////////////////////////////////////////////////////
'//// Creates and assigns a font to a DC
'///////////////////////////////////////////////////////////
Public Sub _
DrawText(hdc As Long, x As Integer, y As Integer, _
            lpszStr As String, _
            Size As Byte, bBold As Boolean, _
            lForeColor As Long, _
            Optional lBackColor = 0, _
            Optional lpszFontName As String = "Arial")

 Dim hFont As Long                                      ' font handle
 Dim hOldFont As Long                                   ' font that will be replaced
 Dim Weight As Long
 
 If bBold Then                                          ' check if bold
  Weight = FW_BOLD
 Else
  Weight = FW_NORMAL
 End If
 
  ' create the font
 hFont = CreateFont(Size, 0, 0, 0, Weight, _
                   False, False, False, _
                   DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, _
                   CLIP_DEFAULT_PRECIS, 0, _
                   DEFAULT_PITCH Or FF_DONTCARE, _
                   lpszFontName)
                   
 hOldFont = SelectObject(hdc, hFont)                    ' select the font into the destination DC
 
 If lBackColor <> -1 Then                               ' check if Transparent
  Call SetBkMode(hdc, TEXT_OPAQUE)
  Call SetBkColor(hdc, lBackColor)
 Else
  Call SetBkMode(hdc, TEXT_TRANSPARENT)
 End If
 
 Call SetTextColor(hdc, lForeColor)
 Call TextOut(hdc, x, y, lpszStr, Len(lpszStr))         ' print text
 
 Call SelectObject(hdc, hOldFont)                       ' restore old font (otherwise it will eat resources EACH frame)
 If DeleteObject(hFont) = 0 Then                        ' finally delete the object ( remove it from heap memory )
    AppendToLog ("ERROR in proc. AssignFont: Error removing font from GDI heap! ")
    Debug.Print "Error removing font from GDI heap!"
 End If

End Sub

