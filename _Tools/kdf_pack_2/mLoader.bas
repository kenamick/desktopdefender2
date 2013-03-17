Attribute VB_Name = "mLoader"
Option Explicit


' Big 10x to Lucky (http://rookscape.com/vbgaming/) _
  for the Great Tutorials

' --- Wave Loader Declarations ---
Private Const RIFF = &H46464952
Private Const WAVE = &H45564157

'File format structure for wave files
Private Type WAVEFORMATEX1
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
        wBitsPerSample As Integer
End Type

'File header structure for wave files.
Private Type WAVEFILEHEADER
    dwRiff As Long
    dwFileSize As Long
    dwWave As Long
    dwFormat As Long
    dwFormatLength As Long
End Type

' --- Bitmap Loader Declarations ---
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Private Const BMP_HEADER = &H4D42 ' BM = 19778
Private Const SRCCOPY = &HCC0020
Private Const DIB_RGB_COLORS = 0

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

Dim bmpHeader As BITMAPFILEHEADER   ' Holds the file header
Dim bmpInfo As BITMAPINFO           ' Holds the bitmap info
Dim bmpData() As Byte               ' Holds the pixel data

' /////////////////////////////////////////////////////////////////
' //// Loads a bitmap from a given position in a binary data packet
' //// hDestDC - is the DC where the bitmap will be copied
' //// lOffset - the position in the binary data packet
' /////////////////////////////////////////////////////////////////
Public Function _
LRLoadBitmap(hDestDC As Long, _
             lpszLibrary As String, _
             lOffset As Long) As Boolean
  
  Erase bmpInfo.bmiColors
  
  Dim cn     As Integer          ' local counter
  Dim rval   As Long
  Dim nFF    As Integer
  Dim nWidth As Integer, nHeight As Integer

  nFF = FreeFile()
  
  ' --- Open File ---
  Open lpszLibrary For Binary Access Read Lock Write As nFF
  ' get fileheader
  Get nFF, lOffset, bmpHeader
  
  ' check for bitamp header
  If (bmpHeader.bfType <> BMP_HEADER) Then
   LRLoadBitmap = False
   Exit Function
  End If
  
  ' get infoheader
  Get nFF, , bmpInfo.bmiHeader
  
  ' --- setup 8 bit images ---
  If (bmpInfo.bmiHeader.biBitCount = 8) Then
    Get nFF, , bmpInfo.bmiColors
    ReDim bmpData(FileSize(bmpInfo.bmiHeader.biWidth, _
                           bmpInfo.bmiHeader.biHeight))
    ' get image
    Get nFF, , bmpData
    
    bmpHeader.bfOffBits = 1078 ' 1024 + 54(header)
    With bmpInfo.bmiHeader
     .biSizeImage = FileSize(bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight)
     .biClrUsed = 0
     .biClrImportant = 0
     .biXPelsPerMeter = 0
     .biYPelsPerMeter = 0
    End With
  
  Else
   ' --- setup 24 bit images ---
  
   ' format palette
   If (bmpInfo.bmiHeader.biClrUsed <> 0) Then
    For cn = 0 To bmpInfo.bmiHeader.biClrUsed - 1
     Get nFF, , bmpInfo.bmiColors(cn).rgbBlue
     Get nFF, , bmpInfo.bmiColors(cn).rgbGreen
     Get nFF, , bmpInfo.bmiColors(cn).rgbRed
     Get nFF, , bmpInfo.bmiColors(cn).rgbReserved
    Next cn
   End If

   ReDim bmpData(bmpInfo.bmiHeader.biSizeImage - 1)
  
   ' get image
   Get nFF, , bmpData

  End If
   
  ' --- Close File ---
  Close nFF
  
  ' now to stretch the (di)bits
  nWidth = bmpInfo.bmiHeader.biWidth
  nHeight = bmpInfo.bmiHeader.biHeight
  
  rval = StretchDIBits(hDestDC, 0, 0, nWidth, nHeight, _
                                0, 0, nWidth, nHeight, _
                                bmpData(0), _
                                bmpInfo, _
                                DIB_RGB_COLORS, _
                                SRCCOPY)
  
 ' return boolean result
 If (rval <> 0) Then LRLoadBitmap = True Else _
                     LRLoadBitmap = False
End Function

Private Function FileSize(lngWidth As Long, lngHeight As Long) As Long

    'Return the size of the image portion of the bitmap
    If lngWidth Mod 4 > 0 Then
        FileSize = ((lngWidth \ 4) + 1) * 4 * lngHeight - 1
    Else
        FileSize = lngWidth * lngHeight - 1
    End If

End Function

