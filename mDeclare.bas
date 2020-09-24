Attribute VB_Name = "mDeclare"
Option Explicit

' Various constants

    Public Const PIFACTOR = 0.0174532
    
' Constants for use with win32 API ...
    
    Public Const IMAGE_BITMAP = 0
    Public Const LR_LOADFROMFILE = &H10
    Public Const LR_CREATEDIBSECTION = &H2000
    Public Const SRCCOPY = &HCC0020

' Types for use with win32 API ...

    ' Wave format type
    Type WAVEFORMATEX
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
        wBitsPerSample As Integer
        cbSize As Integer
    End Type
    
    ' Rectangle type
    Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    
    ' Bitmap descriptor type
    Public Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
    End Type

' Functions for use with win32 API ...

    ' Single Pixel manipulation
    Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
    Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
    
    ' DC manipulation
    Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
    
    ' General GDI Object manipulation
    Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
    Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    
    ' Bitmap manipulation
    Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
    Public Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
    
    ' Various functions
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long
    Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal source As Long, ByVal length As Long)
    
' Types for use with XDemo3D ...
    
    ' Driver type for enumeration of D3D driver
    Public Type tD3DDriver
      DESC    As String                         ' Driver description
      NAME    As String                         ' Driver name
      GUID    As Byte                           ' Unique interface ID for accessing driver
      GUID1   As Byte                           ' ...
      GUID2   As Byte                           ' ...
      GUID3   As Byte                           ' ...
      GUID4   As Byte                           ' ...
      GUID5   As Byte                           ' ...
      GUID6   As Byte                           ' ...
      GUID7   As Byte                           ' ...
      GUID8   As Byte                           ' ...
      GUID9   As Byte                           ' ...
      GUID10  As Byte                           ' ...
      GUID11  As Byte                           ' ...
      GUID12  As Byte                           ' ...
      GUID13  As Byte                           ' ...
      GUID14  As Byte                           ' ...
      GUID15  As Byte                           ' ...
      DEVDESC As D3DDEVICEDESC                  ' Device description for use by D3DRM
      HDW     As Boolean                        ' Device is hardware
      EMU     As Boolean                        ' Device is software-emulated
      RGB     As Boolean                        ' Device has rgb caps
      MONO    As Boolean                        ' Device has mono ramp caps
    End Type
    
    ' Viewport window for scrolling effects windows
    Public Type tDDWindow
        nX As Integer
        nY As Integer
        nDX As Integer
        nDY As Integer
        oDDSurface As IDirectDrawSurface3
        dRenderArea As RECT
    End Type
    
    ' Flying star description
    Public Type tStar
        nX As Long
        nY As Long
        nSpeed As Integer
        nColor As Long
    End Type

    ' Explosion description
    Public Type tExplo
        nX As Long
        nY As Long
        nPhase As Long
        oDSBuffer As IDirectSoundBuffer
    End Type
