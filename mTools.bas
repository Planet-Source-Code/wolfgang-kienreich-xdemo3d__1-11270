Attribute VB_Name = "mTools"
Option Explicit

Public Function EnumDeviceCallback(nGUID As Long, nDDDesc As Long, nDDName As Long, dHALDDDesc As D3DDEVICEDESC, dHELDDDesc As D3DDEVICEDESC, nOptions As Long) As Long

    ' Setup local variables ...
    
        Dim L_nTemp(256) As Byte                      ' Temporary array for name and guid translation
        Dim L_nChar As Byte                           ' Temporary charactar for name translation
        Dim L_nIndex As Long                          ' Variable to run through temp array
    
    ' Process current driver ...
    
        ' Inspect current driver
        With G_dD3DDriver
            
            ' Set driver capabilities...
              
                ' Decide if hardware supports color model and enable HAL or HEL support properly
                If dHALDDDesc.dcmColorModel Then
                  .DEVDESC = dHALDDDesc
                  .HDW = True
                Else
                  .DEVDESC = dHELDDDesc
                End If
                  
                ' Set RGB capability
                .RGB = (.DEVDESC.dcmColorModel = D3DCOLOR_RGB)
                ' Set emulation mode
                .EMU = (Not .HDW)
                ' Set mono ramp mode
                .MONO = (Not .RGB)
              
            ' Decide if driver fits application needs ...
            
                ' Exit without naming/enabling driver if no support for application color depth
                If (.DEVDESC.dwDeviceRenderBitDepth And DDBD_8) = 0 Then
                    EnumDeviceCallback = DDENUMRET_OK
                    Exit Function
                End If
                
                ' Exit without naming/enabling driver if no RGB support
                If Not .RGB Then
                    EnumDeviceCallback = DDENUMRET_OK
                    Exit Function
                End If
    
            ' DRIVER ACCEPTED ...
            
                ' Copy GUID data into temporary array
                CopyMemory VarPtr(L_nTemp(0)), VarPtr(nGUID), 16
                
                ' Set GUID data into driver structure
                .GUID = L_nTemp(0)
                .GUID1 = L_nTemp(1)
                .GUID2 = L_nTemp(2)
                .GUID3 = L_nTemp(3)
                .GUID4 = L_nTemp(4)
                .GUID5 = L_nTemp(5)
                .GUID6 = L_nTemp(6)
                .GUID7 = L_nTemp(7)
                .GUID8 = L_nTemp(8)
                .GUID9 = L_nTemp(9)
                .GUID10 = L_nTemp(10)
                .GUID11 = L_nTemp(11)
                .GUID12 = L_nTemp(12)
                .GUID13 = L_nTemp(13)
                .GUID14 = L_nTemp(14)
                .GUID15 = L_nTemp(15)
                
                ' Copy driver name into temporary array
                CopyMemory VarPtr(L_nTemp(0)), VarPtr(nDDName), 255
                  
                ' Parse name of driver
                For L_nIndex = 0 To 255
                    L_nChar = L_nTemp(L_nIndex)
                    If L_nChar < 32 Then Exit For
                    .NAME = .NAME + Chr(L_nChar)
                Next
                
                ' Copy driver Description into temporary array
                CopyMemory VarPtr(L_nTemp(0)), VarPtr(nDDDesc), 255
                  
                ' Parse description of driver
                For L_nIndex = 0 To 255
                    L_nChar = L_nTemp(L_nIndex)
                    If L_nChar < 32 Then Exit For
                    .DESC = .DESC + Chr(L_nChar)
                Next
        
        End With
            
        ' Return success
        G_bD3DDriverPresent = True
        EnumDeviceCallback = DDENUMRET_CANCEL

End Function

Public Function LoadBitmapIntoDXS(ByVal sFileName As String) As IDirectDrawSurface3

    ' Enable error handling ...
        On Error GoTo E_LoadBitmapIntoDXS
    
    ' Setup local variables ...
        
        Dim L_nBMBitmap As Long               ' Handle on bitmap
        Dim L_nDCBitmap As Long               ' Handle on dc of bitmap
        Dim L_dBitmap As BITMAP               ' Bitmap descriptor
        Dim L_dDXD As DDSURFACEDESC           ' Surface descriptor
        Dim L_nDCDXS As Long                  ' Handle on dc of surface
        Dim L_oDDSTemp As IDirectDrawSurface3 ' Temporary DD surface
    
    ' Load bitmap into surface ...
    
        ' Load bitmap
        L_nBMBitmap = LoadImage(ByVal 0&, sFileName, 0, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
        
        ' Check for validity of bitmap handle
        If L_nBMBitmap < 1 Then
            AppError 0, "Bitmap could not be loaded", "LoadBitmapIntoDXS"
            Exit Function
        End If
        
        ' Get bitmap descriptor
        GetObject L_nBMBitmap, Len(L_dBitmap), L_dBitmap
        
        ' Fill DX surface description
        With L_dDXD
            .dwSize = Len(L_dDXD)
            .dwFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
            .DDSCAPS.dwCaps = DDSCAPS_OFFSCREENPLAIN
            .dwWidth = L_dBitmap.bmWidth
            .dwHeight = L_dBitmap.bmHeight
        End With
        
        ' Create DX surface
        G_oDDInstance.CreateSurface L_dDXD, L_oDDSTemp, Nothing
        
        ' Check surface existance
        If L_oDDSTemp Is Nothing Then
            AppError 0, "Surface could not be created", "LoadBitmapIntoDXS"
            Exit Function
        End If
        
        ' Create API memory DC
        L_nDCBitmap = CreateCompatibleDC(ByVal 0&)
        
        ' Select the bitmap into API memory DC
        SelectObject L_nDCBitmap, L_nBMBitmap
        
        ' Restore DX surface
        L_oDDSTemp.Restore
        
        ' Get DX surface API DC
        L_oDDSTemp.GetDC L_nDCDXS
        
        ' Blit BMP from API DC into DX DC using standard API BitBlt
        StretchBlt L_nDCDXS, 0, 0, L_dDXD.dwWidth, L_dDXD.dwHeight, L_nDCBitmap, 0, 0, L_dBitmap.bmWidth, L_dBitmap.bmHeight, SRCCOPY
        
        ' Cleanup
        L_oDDSTemp.ReleaseDC L_nDCDXS
        DeleteDC L_nDCBitmap
        DeleteObject L_nBMBitmap
        
        ' Return success
        Set LoadBitmapIntoDXS = L_oDDSTemp
        
        ' Cleanup
        Set L_oDDSTemp = Nothing
    
    ' Error handler ...
    
        Exit Function
    
E_LoadBitmapIntoDXS:

        AppError Err.Number, Err.Description, "LoadBitmapIntoDXS"

End Function

Public Function MakeDXSurface(ByVal nWidth As Integer, ByVal nHeight As Integer, Optional bIs3D As Boolean) As IDirectDrawSurface3

    ' Enable error handling ...
        
        On Error GoTo E_MakeDXSurface
    
    ' Setup local variables ...

        Dim L_dDXD As DDSURFACEDESC    ' Variable holding temporary surface description

    ' Create surface ...
    
        ' Fill surface description
        With L_dDXD
           .dwSize = Len(L_dDXD)
           .dwFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
           .DDSCAPS.dwCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY Or IIf(bIs3D, DDSCAPS_3DDEVICE, 0)
           .dwWidth = nWidth
           .dwHeight = nHeight
        End With
    
        ' Create surface from description
        G_oDDInstance.CreateSurface L_dDXD, MakeDXSurface, Nothing
    
        ' Check for existance of surface
        If MakeDXSurface Is Nothing Then
           AppError 0, "Surface could not be created", "MakeDXSurface"
           Exit Function
        End If
    
    ' Error handler ...
        Exit Function
    
E_MakeDXSurface:

        AppError Err.Number, Err.Description, "MakeDXSurface"
    
End Function

Public Function LoadWaveIntoDSB(ByVal sFileName As String) As IDirectSoundBuffer
        
    ' Enable error handling ...
        
        On Error GoTo E_LoadWaveIntoDSB
    
    ' Setup local variables ...
    
        Dim L_dWFX As WAVEFORMATEX      ' Structure holding wave format description
        Dim L_nDataSize As Long         ' Size of audio data
        Dim L_nPosition As Long         ' Current position within wave file
        Dim L_nWaveBytes() As Byte      ' Array holding wave file data
        Dim L_dDSBD As DSBUFFERDESC     ' Structure holding description of DirectSound buffer
        
        Dim L_nPointer1 As Long      ' Pointer to left track data
        Dim L_nLength1 As Long       ' Length of left track data
        
        Dim L_nPointer2 As Long     ' Pointer to right track data
        Dim L_nLength2 As Long      ' Length of right track data
        
    ' Read wave file data into local array ...
        
        ' Set array size to file size
        ReDim L_nWaveBytes(1 To FileLen(sFileName))
        
        ' Load data into array
        Open sFileName For Binary As #1
        Get #1, , L_nWaveBytes
        Close #1
        
    ' Search for format data position ...
                
        ' Start at position 1
        L_nPosition = 1
        
        ' Look for format expression
        Do While Not (Chr(L_nWaveBytes(L_nPosition)) + Chr(L_nWaveBytes(L_nPosition + 1)) + Chr(L_nWaveBytes(L_nPosition + 2)) = "fmt")
            
            L_nPosition = L_nPosition + 1
            
            ' Cancel if no format expression found
            If L_nPosition > UBound(L_nWaveBytes) - 3 Then
                AppError 0, "Invalid file format", "LoadWaveIntoDSB"
            End If
            
        Loop
        
    ' Copy format data to local structure ...
        CopyMemory VarPtr(L_dWFX), VarPtr(L_nWaveBytes(L_nPosition + 8)), Len(L_dWFX)
        
            
    ' Search for audio data position ...
            
        ' Look for data expression
        Do While Not (Chr(L_nWaveBytes(L_nPosition)) + Chr(L_nWaveBytes(L_nPosition + 1)) + Chr(L_nWaveBytes(L_nPosition + 2)) + Chr(L_nWaveBytes(L_nPosition + 3)) = "data")
            
            L_nPosition = L_nPosition + 1
            
            ' Cancel if no data expression found
            If L_nPosition > UBound(L_nWaveBytes) - 4 Then
                AppError 0, "Invalid file format", "LoadWaveIntoDSB"
            End If
            
        Loop
        
        ' Copy data size to local variable
        CopyMemory VarPtr(L_nDataSize), VarPtr(L_nWaveBytes(L_nPosition + 4)), 4
            
    ' Create and fill DirectSound buffer ...
    
        ' Fill description structure to create buffer from
        With L_dDSBD
            .dwSize = Len(L_dDSBD)
            .dwFlags = DSBCAPS_CTRLDEFAULT
            .dwBufferBytes = L_nDataSize
            .lpwfxFormat = VarPtr(L_dWFX)
        End With
        
        ' Create buffer from structure
        G_oDSInstance.CreateSoundBuffer L_dDSBD, LoadWaveIntoDSB, Nothing
        
        ' Check for existance of buffer
        If LoadWaveIntoDSB Is Nothing Then
           AppError 0, "Buffer could not be created", "LoadWaveIntoDSB"
           Exit Function
        End If
        
        ' Lock buffer to access its data
        LoadWaveIntoDSB.Lock 0&, L_nDataSize, L_nPointer1, L_nLength1, L_nPointer2, L_nLength2, 0&
        
        ' Copy left (or only, if mono wave file) track to DirectSound buffer
        CopyMemory L_nPointer1, VarPtr(L_nWaveBytes(L_nPosition + 8)), L_nLength1
        
        ' Copy right track to DirectSound buffer, if exists
        If L_nLength2 <> 0 Then
            CopyMemory L_nPointer2, VarPtr(L_nWaveBytes(L_nPosition + 8 + L_nLength1)), L_nLength2
        End If
        
        ' Unlock buffer
        LoadWaveIntoDSB.Unlock L_nPointer1, L_nLength1, L_nPointer2, L_nLength2
        
    ' Error handler ...
        Exit Function
    
E_LoadWaveIntoDSB:

        AppError Err.Number, Err.Description, "LoadWaveIntoDSB"
    
End Function

Public Function GetDXError(nError As Long) As String

    Select Case nError
        Case DDERR_DCALREADYCREATED
            GetDXError = "A device context (DC) has already been returned for this surface. Only one DC can be retrieved for each surface."
        Case DDERR_DIRECTDRAWALREADYCREATED
            GetDXError = "A IDirectDraw object representing this driver has already been created for this process."
        Case DDERR_EXCEPTION
            GetDXError = "An exception was encountered while performing the requested operation."
        Case DDERR_EXCLUSIVEMODEALREADYSET
            GetDXError = "An attempt was made to set the cooperative level when it was already set to exclusive."
        Case DDERR_HEIGHTALIGN
            GetDXError = "The height of the provided rectangle is not a multiple of the required alignment."
        Case DDERR_HWNDALREADYSET
            GetDXError = "The IDirectDraw cooperative level window handle has already been set. It cannot be reset while the process has surfaces or palettes created."
        Case DDERR_HWNDSUBCLASSED
            GetDXError = "IDirectDraw is prevented from restoring state because the IDirectDraw cooperative level window handle has been subclassed."
        Case DDERR_ALREADYINITIALIZED
            GetDXError = "The object has already been initialized."
        Case DDERR_BLTFASTCANTCLIP
            GetDXError = "A IDirectDrawClipper object is attached to a source surface that has passed into a call to the IDirectDrawSurface2::BltFast method."
        Case DDERR_CANNOTATTACHSURFACE
            GetDXError = "A surface cannot be attached to another requested surface."
        Case DDERR_CANNOTDETACHSURFACE
            GetDXError = "A surface cannot be detached from another requested surface."
        Case DDERR_CANTCREATEDC
            GetDXError = "Windows cannot create any more device contexts (DCs)."
        Case DDERR_CANTDUPLICATE
            GetDXError = "Primary and 3D surfaces, or surfaces that are implicitly created, cannot be duplicated."
        Case DDERR_CANTLOCKSURFACE
            GetDXError = "Access to this surface is refused because an att    empt was made to lock the primary surface without DCI support."
        Case DDERR_CANTPAGELOCK
            GetDXError = "An attempt to page lock a surface failed. Page lock will not work on a display-m    emory surface or an     emulated primary surface."
        Case DDERR_CANTPAGEUNLOCK
            GetDXError = "An attempt to page unlock a surface failed. Page unlock will not work on a display-m    emory surface or an     emulated primary surface."
        Case DDERR_CLIPPERISUSINGHWND
            GetDXError = "An attempt was made to set a clip list for a IDirectDrawClipper object that is already monitoring a window handle."
        Case DDERR_COLORKEYNOTSET
            GetDXError = "No source color key is specified for this operation."
        Case DDERR_CURRENTLYNOTAVAIL
            GetDXError = "No support is currently available."
        Case DDERR_IMPLICITLYCREATED
            GetDXError = "The surface cannot be restored because it is an implicitly created surface."
        Case DDERR_INCOMPATIBLEPRIMARY
            GetDXError = "The primary surface creation request does not match with the existing primary surface."
        Case DDERR_INVALIDCAPS
            GetDXError = "One or more of the capability bits passed to the callback function are incorrect."
        Case DDERR_INVALIDCLIPLIST
            GetDXError = "IDirectDraw does not support the provided clip list."
        Case DDERR_INVALIDDIRECTDRAWGUID
            GetDXError = "The globally unique identifier (GUID) passed to the IDirectDrawCreate function is not a valid IDirectDraw driver identifier."
        Case DDERR_INVALIDMODE
            GetDXError = "IDirectDraw does not support the requested mode."
        Case DDERR_INVALIDOBJECT
            GetDXError = "IDirectDraw received a pointer that was an invalid IDirectDraw object."
        Case DDERR_INVALIDPIXELFORMAT
            GetDXError = "The pixel format was invalid as specified."
        Case DDERR_INVALIDPOSITION
            GetDXError = "The position of the overlay on the destination is no longer legal."
        Case DDERR_INVALIDRECT
            GetDXError = "The provided rectangle was invalid."
        Case DDERR_INVALIDSURFACETYPE
            GetDXError = "The requested operation could not be performed because the surface was of the wrong type."
        Case DDERR_LOCKEDSURFACES
            GetDXError = "One or more surfaces are locked, causing the failure of the requested operation."
        Case DDERR_NO3D
            GetDXError = "No 3D hardware or emulation is present."
        Case DDERR_NOALPHAHW
            GetDXError = "No alpha acceleration hardware is present or available, causing the failure of the requested operation."
        Case DDERR_NOBLTHW
            GetDXError = "No blitter hardware is present."
        Case DDERR_NOCLIPLIST
            GetDXError = "No clip list is available."
        Case DDERR_NOCLIPPERATTACHED
            GetDXError = "No IDirectDrawClipper object is attached to the surface object."
        Case DDERR_NOCOLORCONVHW
            GetDXError = "The operation cannot be carried out because no color-conversion hardware is present or available."
        Case DDERR_NOCOLORKEY
            GetDXError = "The surface does not currently have a color key."
        Case DDERR_NOCOLORKEYHW
            GetDXError = "The operation cannot be carried out because there is no hardware support for the destination color key."
        Case DDERR_NOCOOPERATIVELEVELSET
            GetDXError = "A create function is called without the IDirectDraw2::SetCooperativeLevel method being called."
        Case DDERR_NODC
            GetDXError = "No DC has ever been created for this surface."
        Case DDERR_NODDROPSHW
            GetDXError = "No IDirectDraw raster operation (ROP) hardware is available."
        Case DDERR_NODIRECTDRAWHW
            GetDXError = "Hardware-only IDirectDraw object creation is not possible; the driver does not support any hardware."
        Case DDERR_NODIRECTDRAWSUPPORT
            GetDXError = "IDirectDraw support is not possible with the current display driver."
        Case DDERR_NOEMULATION
            GetDXError = "Software emulation is not available."
        Case DDERR_NOEXCLUSIVEMODE
            GetDXError = "The operation requires the application to have exclusive mode, but the application does not have exclusive mode."
        Case DDERR_NOFLIPHW
            GetDXError = "Flipping visible surfaces is not supported."
        Case DDERR_NOGDI
            GetDXError = "No GDI is present."
        Case DDERR_NOHWND
            GetDXError = "Clipper notification requires a window handle, or no window handle has been previously set as the cooperative level window handle."
        Case DDERR_NOMIPMAPHW
            GetDXError = "The operation cannot be carried out because no mipmap texture mapping hardware is present or available."
        Case DDERR_NOMIRRORHW
            GetDXError = "The operation cannot be carried out because no mirroring hardware is present or available."
        Case DDERR_NOOVERLAYDEST
            GetDXError = "The IDirectDrawSurface2::GetOverlayPosition method is called on an overlay that the IDirectDrawSurface2::UpdateOverlay method has not been called on to establish a destination."
        Case DDERR_NOOVERLAYHW
            GetDXError = "The operation cannot be carried out because no overlay hardware is present or available."
        Case DDERR_NOPALETTEATTACHED
            GetDXError = "No palette object is attached to this surface."
        Case DDERR_NOPALETTEHW
            GetDXError = "There is no hardware support for 16- or 256-color palettes."
        Case DDERR_NORASTEROPHW
            GetDXError = "The operation cannot be carried out because no appropriate raster operation hardware is present or available."
        Case DDERR_NOROTATIONHW
            GetDXError = "The operation cannot be carried out because no rotation hardware is present or available."
        Case DDERR_NOSTRETCHHW
            GetDXError = "The operation cannot be carried out because there is no hardware support for stretching."
        Case DDERR_NOT4BITCOLOR
            GetDXError = "The IDirectDrawSurface object is not using a 4-bit color palette and the requested operation requires a 4-bit color palette."
        Case DDERR_NOT4BITCOLORINDEX
            GetDXError = "The IDirectDrawSurface object is not using a 4-bit color index palette and the requested operation requires a 4-bit color index palette."
        Case DDERR_NOT8BITCOLOR
            GetDXError = "The IDirectDrawSurface object is not using an 8-bit color palette and the requested operation requires an 8-bit color palette."
        Case DDERR_NOTAOVERLAYSURFACE
            GetDXError = "An overlay component is called for a non-overlay surface."
        Case DDERR_NOTEXTUREHW
            GetDXError = "The operation cannot be carried out because no texture-mapping hardware is present or available."
        Case DDERR_NOTFLIPPABLE
            GetDXError = "An attempt has been made to flip a surface that cannot be flipped."
        Case DDERR_NOTFOUND
            GetDXError = "The requested item was not found."
        Case DDERR_NOTLOCKED
            GetDXError = "An attempt is made to unlock a surface that was not locked."
        Case DDERR_NOTPAGELOCKED
            GetDXError = "An attempt is made to page unlock a surface with no outstanding page locks."
        Case DDERR_NOTPALETTIZED
            GetDXError = "The surface being used is not a palette-based surface."
        Case DDERR_NOVSYNCHW
            GetDXError = "The operation cannot be carried out because there is no hardware support for vertical blank synchronized operations."
        Case DDERR_NOZBUFFERHW
            GetDXError = "The operation to create a z-buffer in display memory or to perform a blit using a z-buffer cannot be carried out because there is no hardware support for z-buffers."
        Case DDERR_NOZOVERLAYHW
            GetDXError = "The overlay surfaces cannot be z-layered based on the z-order because the hardware does not support z-ordering of overlays."
        Case DDERR_OUTOFCAPS
            GetDXError = "The hardware needed for the requested operation has already been allocated."
        Case DDERR_OUTOFVIDEOMEMORY
            GetDXError = "IDirectDraw does not have enough display memory to perform the operation."
        Case DDERR_OVERLAYCANTCLIP
            GetDXError = "The hardware does not support clipped overlays."
        Case DDERR_OVERLAYCOLORKEYONLYONEACTIVE
            GetDXError = "An attempt was made to have more than one color key active on an overlay."
        Case DDERR_OVERLAYNOTVISIBLE
            GetDXError = "The IDirectDrawSurface2::GetOverlayPosition method is called on a hidden overlay."
        Case DDERR_PALETTEBUSY
            GetDXError = "Access to this palette is refused because the palette is locked by another thread."
        Case DDERR_PRIMARYSURFACEALREADYEXISTS
            GetDXError = "This process has already created a primary surface."
        Case DDERR_REGIONTOOSMALL
            GetDXError = "The region passed to the IDirectDrawClipper::GetClipList method is too small."
        Case DDERR_SURFACEALREADYATTACHED
            GetDXError = "An attempt was made to attach a surface to another surface to which it is already attached."
        Case DDERR_SURFACEALREADYDEPENDENT
            GetDXError = "An attempt was made to make a surface a dependency of another surface to which it is already dependent."
        Case DDERR_SURFACEBUSY
            GetDXError = "Access to the surface is refused because the surface is locked by another thread."
        Case DDERR_SURFACEISOBSCURED
            GetDXError = "Access to the surface is refused because the surface is obscured."
        Case DDERR_SURFACELOST
            GetDXError = "Access to the surface is refused because the surface memory is gone. The IDirectDrawSurface object representing this surface should have the IDirectDrawSurface2::Restore method called on it."
        Case DDERR_SURFACENOTATTACHED
            GetDXError = "The requested surface is not attached."
        Case DDERR_TOOBIGHEIGHT
            GetDXError = "The height requested by IDirectDraw is too large."
        Case DDERR_TOOBIGSIZE
            GetDXError = "The size requested by IDirectDraw is too large. However, the individual height and width are OK."
        Case DDERR_TOOBIGWIDTH
            GetDXError = "The width requested by IDirectDraw is too large."
        Case DDERR_UNSUPPORTEDFORMAT
            GetDXError = "The FourCC format requested is not supported by IDirectDraw."
        Case DDERR_UNSUPPORTEDMASK
            GetDXError = "The bitmask in the pixel format requested is not supported by IDirectDraw."
        Case DDERR_UNSUPPORTEDMODE
            GetDXError = "The display is currently in an unsupported mode."
        Case DDERR_VERTICALBLANKINPROGRESS
            GetDXError = "A vertical blank is in progress."
        Case DDERR_WASSTILLDRAWING
            GetDXError = "The previous blit operation that is transferring information to or from this surface is incomplete."
        Case DDERR_WRONGMODE
            GetDXError = "This surface cannot be restored because it was created in a different mode."
        Case DDERR_XALIGN
            GetDXError = "The provided rectangle was not horizontally aligned on a required boundary."
        Case Else
            GetDXError = "Unknown Error: Out of memory or invalid parameters passed."
    End Select
    
End Function

