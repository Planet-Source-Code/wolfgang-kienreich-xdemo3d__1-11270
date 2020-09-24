Attribute VB_Name = "mMain"
Option Explicit

' Variables holding DDraw and D3DRM instances ...
Public G_oDDInstance As IDirectDraw                 ' Instance of DirectDraw interface
Public G_oD3DInstance As IDirect3DRM                ' Instance of Direct3DRM interface
Public G_oDSInstance As IDirectSound                ' Instance of DirectSound interface

' Variables for primary D3DRM display system ...
Public G_oD3DDevice As IDirect3DRMDevice2           ' Device to use for Direct3DRM operations
Public G_oD3DViewport As IDirect3DRMViewPort        ' Viewport for Direct3DRM to display results in
Public G_oD3DMasterFrame As IDirect3DRMFrame2       ' Top level frame that contains all other frames
Public G_oD3DCameraFrame As IDirect3DRMFrame2       ' Frame to contain the camera; The viewport is created from this frame
Public G_dD3DDriver As tD3DDriver                   ' Driver for use with Direct3DRM
Public G_bD3DDriverPresent As Boolean               ' Flag holding presence of driver (equals driver enumeration success)
Public G_dCamPosLookup(359) As D3DVECTOR            ' Lookup table of position values for camera
Public G_nCamPosCurrent As Integer                  ' Current position of camera according to lookup table
Public G_oDDSurfaceStatus As IDirectDrawSurface3    ' Surface holding text for status of d3drm

' Variables for sound system ...
Public G_oDSBufferMusic As IDirectSoundBuffer       ' Buffer holding constant background music

' Variables for DirectDraw blit system ...
Public G_oDDPrimary As IDirectDrawSurface3          ' Primary DirectDraw surface that is displayed on the form
Public G_oDDBackbuffer As IDirectDrawSurface3       ' Backbuffer DirectDraw surface that is flipped onto the primary
Public G_dDDWindow(2) As tDDWindow                  ' Buffers holding windows for effects

' Variables for rock and lava surface ...
Public G_oD3DTextureGround As IDirect3DRMTexture2  ' Texture for ground terrain
Public G_oD3DMaterialGround As IDirect3DRMMaterial ' Material for ground to add specularity
Public G_oD3DTextureLava As IDirect3DRMTexture2    ' Texture for animated Lava
Public G_oDDSurfaceLava As IDirectDrawSurface3     ' Surface holding current animated Lava
Public G_oDDResourceLava As IDirectDrawSurface3    ' Surface holding original Lava bitmap
Public G_oD3DMaterialLava As IDirect3DRMMaterial   ' Material for lava to make lava emissive

' Variables for rotor animation ...
Public G_oD3DRotorFrame As IDirect3DRMFrame2       ' Frame to hold rotor object
Public G_oD3DTextureRotor As IDirect3DRMTexture2   ' Texture for rotor object
Public G_oD3DMaterialRotor As IDirect3DRMMaterial  ' Material for rotor object

' Variables for flame decal ...
Public G_oDDResourceFlame As IDirectDrawSurface3   ' Surface containing images for flame animation
Public G_oD3DLightFlame1 As IDirect3DRMLight       ' Light for flame to illuminate surroundings
Public G_oD3DFrameFlame1 As IDirect3DRMFrame       ' Frame to contain light for flame
Public G_oDDSurfaceFlame1 As IDirectDrawSurface3   ' Surface containing current state of flame animation
Public G_oD3DTextureFlame1 As IDirect3DRMTexture2  ' Texture to contain decal
Public G_oD3DLightFlame2 As IDirect3DRMLight       ' Light for flame to illuminate surroundings
Public G_oD3DFrameFlame2 As IDirect3DRMFrame       ' Frame to contain light for flame
Public G_oDDSurfaceFlame2 As IDirectDrawSurface3   ' Surface containing current state of flame animation
Public G_oD3DTextureFlame2 As IDirect3DRMTexture2  ' Texture to contain decal

' Variables for mirror effect ...
Public G_oD3DTextureMirror As IDirect3DRMTexture2  ' Texture for mirror effect
Public G_oDDSurfaceMirror As IDirectDrawSurface3   ' Surface for mirror effect
Public G_oD3DViewportMirror As IDirect3DRMViewPort ' Viewport for mirror effect
Public G_oD3DDeviceMirror As IDirect3DRMDevice     ' Device for mirror effect
Public G_oD3DFrameMirror As IDirect3DRMFrame       ' Frame for mirror effect
Public G_oD3DMaterialMirror As IDirect3DRMMaterial ' Material for mirror effect

' Variables for flying text ...
Public G_bFontData(255, 34) As Boolean             ' Array holding character data
Public G_sDisplayText As String                    ' Text to display using scrolling characters
Public G_nCharScrollPos As Integer                 ' Current scroll offset of text
Public G_oDDSurfaceChars As IDirectDrawSurface3    ' Surface holding characters
 
' Variables for background animation ...
Public G_dStar(1999) As tStar                      ' Array holding data of moving stars
Public G_dExplo(14) As tExplo                      ' Array holding data on explosions
Public G_oDDSurfaceExplo As IDirectDrawSurface3    ' Surface holding explosion animations

' Various variables ...
Public G_nFrameCount As Long                       ' Global framecounter
Public G_nFrameAvg As Double                       ' Global average frames per second

' APPERROR: Reports application errors and terminates application properly
Public Sub AppError(nNumber As Long, sText As String, sSource As String)

    ' Enable error handling
    On Error Resume Next
    
    ' Cleanup
    Call AppTerminate
    
    ' Display error
    MsgBox "ERROR: " & IIf(InStr(1, UCase(sText), "AUTOM") > 0, "DirectX reports '" & GetDXError(nNumber) & "'", " Application reports '" & sText & "'") & vbCrLf & "SOURCE: " & sSource, vbCritical + vbOKOnly, "XDEMO3D"
    
    ' Terminate program
    End
    
End Sub

Public Sub AppInitialize()

    ' Enable error handling
        On Error GoTo E_AppInitialize

    ' Setup local variables...
    
        Dim L_dDDSD As DDSURFACEDESC           ' Utility surface description
        Dim L_dDDSC As DDSCAPS                 ' Utility display capabilities description
        Dim L_oD3DIM As IDirect3D2             ' Utility Direct3DIM interface for retrieving drivers
        Dim L_dDDCK As DDCOLORKEY              ' Color key for applying to various surfaces
        
    ' Initialize scrolling text ...
    
        G_sDisplayText = "             welcome to xdemo3d ... explore the world of directx ... explore the world of visual basic ... written by wolfgang kienreich in september 1998 ... contact me at wolfgang.kienreich@dige.com ... thanx to patrice scribe for the great directx type library ... again vb rules ... feel free to spread this demo ... mail me if you've got some interesting vb stuff, or experience any problems with vb and directx ...          "
        
    ' Initialize DirectDraw interface instance ...
    
        ' Create DirectDraw instance
        DirectDrawCreate ByVal 0&, G_oDDInstance, Nothing

        ' Check instance existance, terminate if missing
        If G_oDDInstance Is Nothing Then
           AppError 0, "Could not create DirectDraw instance", "AppInitialize"
           Exit Sub
        End If
         
        ' Set cooperation mode of DirectX
        
        G_oDDInstance.SetCooperativeLevel fMain.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
        
        ' Set display mode
        G_oDDInstance.SetDisplayMode 640, 480, 16
           
    ' Initialize primary surface ...
    
        ' Initialize primary surface description
        With L_dDDSD
            ' Get Structure size
            .dwSize = Len(L_dDDSD)
            ' Structure uses Surface Caps and count of BackBuffers
            .dwFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
            ' Structure describes a flippable (buffered) surface
            .DDSCAPS.dwCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX Or DDSCAPS_SYSTEMMEMORY
            ' Structure uses one BackBuffer
            .dwBackBufferCount = 1
         End With
    
        ' Create primary surface from structure
        G_oDDInstance.CreateSurface L_dDDSD, G_oDDPrimary, Nothing
    
        ' Check primary existance, terminate if missing
        If G_oDDPrimary Is Nothing Then
           AppError 0, "Could not create primary surface", "AppInitialize"
           Exit Sub
        End If
    
    ' Initialize backbuffer from primary ...
    
        ' Set surface description to backbuffer creation
        L_dDDSD.dwFlags = DDSD_CAPS
        L_dDDSD.DDSCAPS.dwCaps = DDSCAPS_BACKBUFFER
        
        ' Create backbuffer from frontbuffer
        G_oDDPrimary.GetAttachedSurface L_dDDSD.DDSCAPS, G_oDDBackbuffer
        
        ' Check backbuffer existance, terminate if missing
        If G_oDDBackbuffer Is Nothing Then
           AppError 0, "Could not create backbuffer", "AppInitialize"
           Exit Sub
        End If
        

    ' Initialize sound system ...
    
        ' Create an instance of DirectSound interface
        DirectSoundCreate ByVal 0&, G_oDSInstance, Nothing
            
        ' Check instance existance, terminate if missing
        If G_oDSInstance Is Nothing Then
           AppError 0, "Could not create DirectSound instance", "AppInitialize"
           Exit Sub
        End If
                    
        ' Set sound system cooperative level
        G_oDSInstance.SetCooperativeLevel fMain.hWnd, DSSCL_NORMAL
        
    ' Initialize background music
        Set G_oDSBufferMusic = LoadWaveIntoDSB(App.Path + "\music.wav")
        G_oDSBufferMusic.Play ByVal 0&, ByVal 0&, DSBPLAY_LOOPING
        
    ' Initialize windows for displaying various effects
        Call CreateWindows
            
    ' Initialize Direct3DRM interface instance ...
    
        ' Create Direct3DRM instance
        Direct3DRMCreate G_oD3DInstance
    
        ' Check instance existance, terminate if missing
        If G_oD3DInstance Is Nothing Then
           AppError 0, "Could not create D3DRM instance", "AppInitialize"
           Exit Sub
        End If
    
    ' Initialize Direct3DRM driver ...
    
        ' Get a Direct3D immediate object from the existing DirectDraw object
        Set L_oD3DIM = G_oDDInstance
    
        ' Set error handler to local for enumeration only
        On Error Resume Next
        
        ' Start the callback that does the driver enumeration
        L_oD3DIM.EnumDevices AddressOf EnumDeviceCallback, 0
    
        ' Catch any error resulting from the enumeration and terminate
        If Err.Number > 0 Then
           AppError Err.Number, Err.Description, "AppInitialize"
           Exit Sub
        End If
    
        ' Reset error handler to default
        On Error GoTo E_AppInitialize
        
        ' Reset Direct3D immediate object
        Set L_oD3DIM = Nothing
        
        ' Check if a convenient device driver has been found, terminate if no driver available
        If Not G_bD3DDriverPresent Then
           AppError 0, "Could not detect Direct3D device driver", "AppInitialize"
           Exit Sub
        End If
        
    ' Initialize D3DRM display system
    
        ' Create a D3DRM device from the 3D buffer
        G_oD3DInstance.CreateDeviceFromSurface G_dD3DDriver.GUID, G_oDDInstance, G_dDDWindow(0).oDDSurface, G_oD3DDevice
        
        ' Check device existance, terminate if missing
        If G_oD3DDevice Is Nothing Then
           AppError 0, "Could not create D3DRM device", "AppInitialize"
           Exit Sub
        End If
    
        ' Set D3DRM device quality
        G_oD3DDevice.SetQuality D3DRMLIGHT_ON Or D3DRMFILL_SOLID Or D3DRMSHADE_PHONG
        
        ' Create the master frame containing all other frames
        G_oD3DInstance.CreateFrame Nothing, G_oD3DMasterFrame
        
        ' Create the camera frame containing the primary camera
        G_oD3DInstance.CreateFrame G_oD3DMasterFrame, G_oD3DCameraFrame
        
        ' Create a D3D viewport from the device, using the camera frame for output
        G_oD3DInstance.CreateViewport G_oD3DDevice, G_oD3DCameraFrame, 0, 0, 280, 200, G_oD3DViewport
            
        ' Check viewport existance, terminate if missing
        If G_oD3DViewport Is Nothing Then
           AppError 0, "Could not create D3DRM viewport", "AppInitialize"
           Exit Sub
        End If
            
        
    ' Initialize scene and display settings ...
            
        Call CreateScene
        
        ' Create character fontset and objects
        Call CreateChars
        
        ' Create textured and animated ground
        Call CreateGround
        
        ' Create mirror effect
        Call CreateMirror
        
        ' Create decal fire
        Call CreateFlames
        
        ' Create rotor object
        Call CreateRotor
        
        ' Create background animation
        Call CreateBack
        
        ' Error handling ...
        
        Exit Sub

E_AppInitialize:

        AppError Err.Number, Err.Description, "AppInitialize"
        
End Sub

Public Sub AppLoop()

    ' Enable error handling
        On Error GoTo E_AppLoop

    ' Setup local variables...
        Dim L_nNextFrameTime As Long        ' Timer used to time frames to a minimum duration
        Dim L_nFrameCount As Long           ' Frame counter used for calculating average framerate
        Dim L_nNextSecond As Long           ' TimeGetTime value above which next second begins
        Dim L_nCurrentTime As Long          ' Time at start of frame, to avoid multiple calls of TimeGetTime
        
        Dim L_dRenderArea As RECT           ' Rectangle to describe render area for blitting
        Dim L_dDDBLTFX As DDBLTFX           ' Holds F/X settings for blitting
        
    ' Preparations for master loop
    
        ' Prepare BLTFX structure for color fill blit to clear backbuffer
        With L_dDDBLTFX
            .dwSize = Len(L_dDDBLTFX)
            .dwFillColor = 0
        End With
        
    ' Master loop controlling application behavior...
        
        Do
            
            ' Do frame timing and statistics ...
            
                ' Increase global frame counter
                G_nFrameCount = G_nFrameCount + 1
                
                ' Get frame start time
                L_nCurrentTime = timeGetTime
                
                ' Increase frame count for avg frametime calculation
                L_nFrameCount = L_nFrameCount + 1
                
                ' Protocol frame time: Count frames and write out average frame count every second
                If L_nNextSecond < L_nCurrentTime Then
                    G_nFrameAvg = (G_nFrameAvg + L_nFrameCount) / 2
                    L_nNextSecond = L_nCurrentTime + 1000
                    L_nFrameCount = 0
                End If
            
                ' Prepare timing: Set next frame time to current time plus minimum frame duration (15fps , makes for ~60ms)
                L_nNextFrameTime = L_nCurrentTime + 50
            
                ' Query user input
                DoEvents
                        
            ' Clear backbuffer ...
                
                ' FX-Blit filling background with black
                With L_dRenderArea
                    .Top = 0
                    .Left = 0
                    .Bottom = 480
                    .Right = 640
                End With
                G_oDDBackbuffer.Blt L_dRenderArea, ByVal Nothing, ByVal 0&, DDBLT_COLORFILL, L_dDDBLTFX
            
            ' Do updating for background animation
                Call UpdateBack
                
            ' Do updating for animated Lava ...
                 If G_nFrameCount Mod 2 = 0 Then Call UpdateGround
                
            ' Update D3DRM only if within active time segment ...
            
                If (G_nFrameCount Mod 150) < 100 Then
                
                    ' Do flame decal updating ...
                         Call UpdateFlames
                        
                    ' Do mirror updating ...
                         If G_nFrameCount Mod 2 = 0 Then Call UpdateMirror
                        
                    ' Do updating for D3DRM scene ...
                         Call UpdateScene
                         
                End If
            
                ' Update status text describing current state of D3DRM window
                With L_dRenderArea
                    .Top = IIf((G_nFrameCount Mod 150) < 100, 0, 1) * 10
                    .Left = 0
                    .Right = 110
                    .Bottom = .Top + 10
                End With
                G_dDDWindow(0).oDDSurface.BltFast 165, 185, G_oDDSurfaceStatus, L_dRenderArea, DDBLTFAST_NOCOLORKEY
                
            ' Update display system ...
               Call UpdateWindows
                                
            ' Do updating of moving characters ...
               Call UpdateChars
                
            ' Flip DirectX buffers...
                G_oDDPrimary.Flip Nothing, 0
                
            ' Do timing: Loop until minimum time per frame reached ...
            Do
            Loop Until timeGetTime > L_nNextFrameTime
        
        Loop
        
    ' Error handling ...
    
        Exit Sub

E_AppLoop:

        ' Resume to ignore the weird math errors Direct3DRM reports from time to time
        Resume Next

End Sub

Public Sub AppTerminate()

    ' Enable error handling...
        On Error GoTo E_AppTerminate

    ' Setup local variables ...
        Dim L_nRun As Integer             ' Variable to run through various array data
    
    ' Return control from DirectX to windows ...

        ' Restore old resolution and depth
        G_oDDInstance.RestoreDisplayMode
    
        ' Return control to windows
        G_oDDInstance.SetCooperativeLevel fMain.hWnd, DDSCL_NORMAL
        
    ' Reset DirectX objects ...
        
        ' D3DRM Flame animation ...
            
            Set G_oD3DLightFlame1 = Nothing
            Set G_oD3DTextureFlame1 = Nothing
            Set G_oDDSurfaceFlame1 = Nothing
            Set G_oD3DFrameFlame1 = Nothing
            
            Set G_oD3DLightFlame2 = Nothing
            Set G_oD3DTextureFlame2 = Nothing
            Set G_oDDSurfaceFlame2 = Nothing
            Set G_oD3DFrameFlame2 = Nothing
            
            Set G_oDDResourceFlame = Nothing
        
        ' D3DRM Lava animation ...
        
            Set G_oD3DTextureLava = Nothing
            Set G_oDDSurfaceLava = Nothing
            Set G_oDDResourceLava = Nothing
            Set G_oD3DMaterialLava = Nothing
            Set G_oD3DMaterialGround = Nothing
        
        ' D3DRM Mirror animation ...
        
            Set G_oD3DTextureMirror = Nothing
            Set G_oD3DFrameMirror = Nothing
                
        ' D3DRM Rotor animation ...
        
            Set G_oD3DRotorFrame = Nothing
            Set G_oD3DTextureRotor = Nothing
            Set G_oD3DMaterialRotor = Nothing
        
        ' D3DRM display system ...
        
            Set G_oD3DCameraFrame = Nothing
            Set G_oD3DMasterFrame = Nothing
            Set G_oD3DViewport = Nothing
            Set G_oD3DDevice = Nothing
            
            Set G_oD3DViewportMirror = Nothing
            Set G_oD3DDeviceMirror = Nothing
            Set G_oDDSurfaceMirror = Nothing
            Set G_oD3DMaterialMirror = Nothing
        
        ' DD Display system ...
        
            Set G_oDDBackbuffer = Nothing
            Set G_oDDPrimary = Nothing
        
        ' DD Character animation ...
        
            Set G_oDDSurfaceChars = Nothing
                
        ' DD Windows ...
        
            For L_nRun = 0 To 4
                Set G_dDDWindow(L_nRun).oDDSurface = Nothing
            Next
                
        ' DD status text for D3DRM window...
        
            Set G_oDDSurfaceStatus = Nothing
        
        ' DD Explosions ...
            
            Set G_oDDSurfaceExplo = Nothing
                
        ' DirectSound ...
        
            G_oDSBufferMusic.Stop
            Set G_oDSBufferMusic = Nothing
            
            For L_nRun = 0 To 7
                If Not G_dExplo(L_nRun).oDSBuffer Is Nothing Then
                    G_dExplo(L_nRun).oDSBuffer.Stop
                    Set G_dExplo(L_nRun).oDSBuffer = Nothing
                End If
            Next
        
        ' DirectX interfaces ...
        
            Set G_oDSInstance = Nothing
            Set G_oDDInstance = Nothing
            Set G_oD3DInstance = Nothing
        
    ' Error handling ...
        
        Exit Sub

E_AppTerminate:

        ' Resume to ensure that all objects available are cleaned up
        Resume Next

End Sub




Private Sub CreateScene()

    ' Enable error handling...
        On Error GoTo E_CreateScene

    ' Setup local variables ...
        Dim L_oD3DLight As IDirect3DRMLight     ' Variable for light creating
        Dim L_nRun As Single                    ' Variable to run through arrays
        Dim L_dDDCK As DDCOLORKEY               ' Color key for making status display transparent
        
    ' Initialize scenario settings ...
    
        ' Create position lookup table for camera
        For L_nRun = 0 To 179
            With G_dCamPosLookup(L_nRun)
                .z = 5
                .X = 11 + Sin((L_nRun * 2) * PIFACTOR) * 7.8
                .Y = 10.5 + Cos((L_nRun * 2) * PIFACTOR) * 7.8
            End With
        Next
        
        ' Set the projection model and properties for the viewport
        With G_oD3DViewport
            .SetProjection D3DRMPROJECT_PERSPECTIVE
            .SetBack 20
            .SetFront 1
            .SetUniformScaling 1
        End With
        
        ' Set the scene properties (fog, backcolor)
        With G_oD3DMasterFrame
            .SetSceneBackgroundRGB 0, 0, 0
            .SetSceneFogColor D3DRMCreateColorRGB(0.2, 0.2, 0.3)
            .SetSceneFogMode D3DRMFOG_EXPONENTIAL
            .SetSceneFogParams 1, 16, 0.1
            .SetSceneFogEnable 1
        End With
        
        ' Create ambient light
        G_oD3DInstance.CreateLightRGB D3DRMLIGHT_AMBIENT, 0.3, 0.3, 0.3, L_oD3DLight
        G_oD3DMasterFrame.AddLight L_oD3DLight
        Set L_oD3DLight = Nothing

        ' Load text for status display
        Set G_oDDSurfaceStatus = LoadBitmapIntoDXS(App.Path + "\text.bmp")
                
        ' Make text transparent
        L_dDDCK.dwColorSpaceHighValue = 0
        L_dDDCK.dwColorSpaceLowValue = 0
        G_oDDSurfaceStatus.SetColorKey DDCKEY_SRCBLT, L_dDDCK
        
    ' Error handling ...
    
        Exit Sub

E_CreateScene:

        AppError Err.Number, Err.Description, "CreateScene"
        Exit Sub

End Sub

Private Sub CreateRotor()

    ' Enable error handling...
        On Error GoTo E_CreateRotor

    ' Setup local variables ...
        
        Dim L_oD3DMeshbuilder As IDirect3DRMMeshBuilder2    ' Holds and loads the mesh for the rotor
        Dim L_oD3DWrap As IDirect3DRMWrap                   ' Wrap for calculating texture coordinates
        
    ' Create and load rotor mesh...
        
        ' Create meshbuilder to hold rotor
        G_oD3DInstance.CreateMeshBuilder L_oD3DMeshbuilder
        
        ' Load rotor from xfile
        L_oD3DMeshbuilder.Load App.Path + "\rotor.x", 0, 0, 0, 0
        
    ' Create and load rotor texture...
    
        ' Load rotor texture
        G_oD3DInstance.LoadTexture App.Path + "\rotor.bmp", G_oD3DTextureRotor
        
        ' Set texture to transparent
        G_oD3DTextureRotor.SetDecalTransparentColor 0
        G_oD3DTextureRotor.SetDecalTransparency 1
        
        ' Calculate texture coordinates for spherical wrapping
        G_oD3DInstance.CreateWrap D3DRMWRAP_SPHERE, Nothing, 0, 0, 0, 0, 0, 1, 0, 1, 0, 0, 0, 1, 1, L_oD3DWrap
        L_oD3DWrap.Apply L_oD3DMeshbuilder
        Set L_oD3DWrap = Nothing
        
        ' Apply texture to mesh
        L_oD3DMeshbuilder.SetTexture G_oD3DTextureRotor
                        
        ' Create and apply material for rotor
        G_oD3DInstance.CreateMaterial 1, G_oD3DMaterialRotor
        G_oD3DMaterialRotor.SetSpecular 0.4, 0.4, 0.4
        L_oD3DMeshbuilder.SetMaterial G_oD3DMaterialRotor
        
    ' Prepare frame for rotor ...
    
        ' Create and locate frame to hold rotor
        G_oD3DInstance.CreateFrame G_oD3DMasterFrame, G_oD3DRotorFrame
        G_oD3DRotorFrame.SetPosition Nothing, 11, 11, 1.5
        G_oD3DRotorFrame.SetRotation Nothing, 0, 0, 1, -0.1
        
        ' Add rotor mesh to frame
        G_oD3DRotorFrame.AddVisual L_oD3DMeshbuilder
        
    ' Cleanup ...
    
        ' Clean up mesh
        Set L_oD3DMeshbuilder = Nothing
        
    ' Error handling ...
    
        Exit Sub

E_CreateRotor:

        AppError Err.Number, Err.Description, "CreateRotor"
        Exit Sub

End Sub

Private Sub CreateBack()

    ' Enable error handling...
        On Error GoTo E_CreateBack

    ' Setup local variables ...
        Dim L_nRun As Integer       ' Variable to run through array
        Dim L_dDDCK As DDCOLORKEY   ' Variable holding transparency key for explosion surface
        
    ' Create background stars ...
        
        For L_nRun = 0 To 1999
            With G_dStar(L_nRun)
                
                ' Set position
                .nX = Int(Rnd * 620) + 10
                .nY = Int(Rnd * 460) + 10
                                
                ' Set speed and color (the further "back", the slower and darker)
                Select Case Int(Rnd * 9) + 1
                
                    Case 1
                        .nSpeed = 3
                        .nColor = 65535
                    Case 2, 3, 4
                        .nSpeed = 2
                        .nColor = 14799 ' 31711
                    Case 5, 6, 7, 8, 9
                        .nSpeed = 1
                        .nColor = 6343 '14799
                        
                End Select
                
            End With
        Next
        
    ' Create background explosions ...
    
        ' Load surface holding explosion bitmaps
        Set G_oDDSurfaceExplo = LoadBitmapIntoDXS(App.Path + "\explo.bmp")
        
        ' Make surface transparent
        L_dDDCK.dwColorSpaceHighValue = 0
        L_dDDCK.dwColorSpaceLowValue = 0
        G_oDDSurfaceExplo.SetColorKey DDCKEY_SRCBLT, L_dDDCK
        
        ' Setup position, speed and sound for individual explosions
        For L_nRun = 0 To 14
            With G_dExplo(L_nRun)
            
                ' Initialize position and speed
                .nX = Int(Rnd * 600) + 5
                .nY = Int(Rnd * 460) + 10
                .nPhase = Int(Rnd * 15)
                
                ' Initialize sound for the first 8 explosions
                If L_nRun < 8 Then Set .oDSBuffer = LoadWaveIntoDSB(App.Path + "\explo.wav")
                
            End With
        Next
    
    ' Error handling ...
    
        Exit Sub

E_CreateBack:

        AppError Err.Number, Err.Description, "CreateBack"
        Exit Sub

End Sub

Private Sub CreateGround()

    ' Enable error handling
        On Error GoTo E_CreateGround

    ' Setup local variables ...
        Dim L_oD3DFace As IDirect3DRMFace                   ' Face to be added to meshbuilder
        Dim L_oD3DMeshbuilder As IDirect3DRMMeshBuilder2    ' Meshbuilder to hold created faces
        Dim L_nRunColumn As Single                                ' Variable to run through x face coordinates
        Dim L_nRunRow As Single                                ' Variable to run through y face coordinates
        
    ' Create ground ...
        
        ' Initialize meshbuilder
        G_oD3DInstance.CreateMeshBuilder L_oD3DMeshbuilder
        
        ' Load ground texture from file
        G_oD3DInstance.LoadTexture App.Path + "\ground.bmp", G_oD3DTextureGround
        
        ' Create Lava texture from surface, load Lava texture into surface
        Set G_oDDResourceLava = LoadBitmapIntoDXS(App.Path + "\lava.bmp")
        Set G_oDDSurfaceLava = MakeDXSurface(32, 32)
        G_oD3DInstance.CreateTextureFromSurface G_oDDSurfaceLava, G_oD3DTextureLava
        
        ' Create emissive material for lava
        G_oD3DInstance.CreateMaterial 1, G_oD3DMaterialLava
        G_oD3DMaterialLava.SetEmissive 0.5, 0.2, 0.1
        G_oD3DMaterialLava.SetSpecular 0.5, 0.5, 0.5
                
        ' Create specular material for ground
        G_oD3DInstance.CreateMaterial 1, G_oD3DMaterialGround
        G_oD3DMaterialGround.SetSpecular 0.5, 0.5, 0.5
        
        ' Create ground faces
        For L_nRunColumn = 8 To 13
            For L_nRunRow = 8 To 12
            
                ' Create face: Set vertices and texture coordinates
                G_oD3DInstance.CreateFace L_oD3DFace
                With L_oD3DFace
                    .AddVertex L_nRunColumn, L_nRunRow, 1
                    .AddVertex L_nRunColumn + 1, L_nRunRow, 1
                    .AddVertex L_nRunColumn + 1, L_nRunRow + 1, 1
                    .AddVertex L_nRunColumn, L_nRunRow + 1, 1
                    .SetTextureCoordinates 0, L_nRunColumn, L_nRunRow
                    .SetTextureCoordinates 1, L_nRunColumn + 1, L_nRunRow
                    .SetTextureCoordinates 2, L_nRunColumn + 1, L_nRunRow + 1
                    .SetTextureCoordinates 3, L_nRunColumn, L_nRunRow + 1
                End With
                
                ' Set face texture: Lava or Ground ,depending on position
                If (L_nRunRow < 10) Then
                    L_oD3DFace.SetTexture G_oD3DTextureLava
                    L_oD3DFace.SetMaterial G_oD3DMaterialLava
                Else
                    L_oD3DFace.SetTexture G_oD3DTextureGround
                    L_oD3DFace.SetMaterial G_oD3DMaterialGround
                End If
                
                ' Add face data to meshbuilder
                L_oD3DMeshbuilder.AddFace L_oD3DFace
            
                ' Reset face
                Set L_oD3DFace = Nothing
                
            Next
        Next
    
        ' Generate lighting normals and compact mesh
        L_oD3DMeshbuilder.GenerateNormals2 1, 3

        ' Add created mesh to frame
        G_oD3DMasterFrame.AddVisual L_oD3DMeshbuilder

        ' Clean up all DirectX objects
        Set L_oD3DMeshbuilder = Nothing
    
    ' Error handling ...
    
        Exit Sub

E_CreateGround:

        AppError Err.Number, Err.Description, "CreateGround"
        Exit Sub
    
End Sub

Private Sub CreateWindows()

    ' Enable error handling...
        On Error GoTo E_CreateWindows

    ' Setup local variables ...
    Dim L_dDDCK As DDCOLORKEY
    
    L_dDDCK.dwColorSpaceHighValue = 0
    L_dDDCK.dwColorSpaceLowValue = 0
    
    ' Initialize threed buffer surface ...
    
        With G_dDDWindow(0)
        
            ' Create surface capable of being a 3d device
            Set .oDDSurface = MakeDXSurface(280, 200, True)
   
             ' Set color key to enable transparent blits
            .oDDSurface.SetColorKey DDCKEY_SRCBLT, L_dDDCK
            
            ' Set render area
            .dRenderArea.Top = 0
            .dRenderArea.Left = 0
            .dRenderArea.Bottom = 200
            .dRenderArea.Right = 280
            
            ' Set initial position and motion
            .nX = 300
            .nY = 100
            .nDX = 2
            .nDY = -2
            
        End With
    
    ' Initialize window holding scrolling font (no surface, position data only) ...
    
        With G_dDDWindow(1)
        
            ' Set initial position and motion
            .nX = 60
            .nY = 400
            .nDX = 0
            .nDY = -1
            
            ' Set render area
            .dRenderArea.Top = 0
            .dRenderArea.Left = 0
            .dRenderArea.Bottom = 60
            .dRenderArea.Right = 500
            
        End With
        
                
    ' Initialize window holding bumping logo (surface is being loaded, not defined here) ...
         With G_dDDWindow(2)
         
            ' Load surface image
            Set .oDDSurface = LoadBitmapIntoDXS(App.Path + "\logo.bmp")
        
            ' Set color key to enable transparent blits
            .oDDSurface.SetColorKey DDCKEY_SRCBLT, L_dDDCK
         
            ' Set initial position and motion
            .nX = 100
            .nY = 100
            .nDY = 2
            .nDX = 2
            
            ' Set render area
            .dRenderArea.Top = 0
            .dRenderArea.Left = 0
            .dRenderArea.Bottom = 100
            .dRenderArea.Right = 200
            
        End With
    

    ' Error handling ...
    
        Exit Sub

E_CreateWindows:
    
    AppError Err.Number, Err.Description, "CreateWindows"
    
End Sub

Private Sub CreateChars()

    ' Enable error handling...
        On Error GoTo E_CreateChars

    ' Setup local variables ...
        
        Dim L_nRunChar As Integer         ' Variable to run through all chars to be created
        Dim L_sInfo As String             ' String holding first line of def, which assigns characters to charset positions
        Dim L_sLine(6) As String          ' Stings to hold character definition
        Dim L_nCharCount As Integer       ' Number of chars to be created
        Dim L_nRunColumn As Integer       ' Variable to run through the bits within a character
        Dim L_nRunRow As Integer          ' Variable to run through the bits within a character
        Dim L_dDDCK As DDCOLORKEY         ' Colorkey for enabling transparent blits from character surface
        
    ' Load character set from definition file ...
        
        ' Open definition file
        Open App.Path + "\font.def" For Input As #1
        
        ' Read first line, which defines which characters are present in the file
        Input #1, L_sInfo
        L_nCharCount = Len(L_sInfo)
        
        ' Read the seven (scan)lines defining the character set
        For L_nRunRow = 0 To 6
            Input #1, L_sLine(L_nRunRow)
        Next
        
        ' Close the definition file
        Close #1
            
    ' Create character array data from loaded data ...
        
        ' Run through all characters to be created
        For L_nRunChar = 1 To L_nCharCount
                
            ' Run through all bits within the definition of the current char
            For L_nRunColumn = 0 To 4
                For L_nRunRow = 0 To 6
                    ' Set element at current position
                    G_bFontData(Asc(Mid(L_sInfo, L_nRunChar, 1)), L_nRunRow * 5 + L_nRunColumn) = Not (Mid(L_sLine(L_nRunRow), 6 * (L_nRunChar - 1) + L_nRunColumn + 1, 1) = " ")
                Next
            Next
            
        Next
       
    ' Initialize character position
        G_nCharScrollPos = 59
    
    ' Initialize character surface...
        
            ' Create surface
            Set G_oDDSurfaceChars = MakeDXSurface(520, 47)
        
            ' Set color key to enable transparent blits
            L_dDDCK.dwColorSpaceLowValue = 0
            L_dDDCK.dwColorSpaceHighValue = 0
            G_oDDSurfaceChars.SetColorKey DDCKEY_SRCBLT, L_dDDCK
            
    ' Error handling ...
    
        Exit Sub

E_CreateChars:
    
        AppError Err.Number, Err.Description, "CreateChars"
    
End Sub
    
Private Sub CreateFlames()


    ' Enable error handling...
        On Error GoTo E_CreateFlames

    ' Setup local variables ...
    
    ' Create flame decals ...
        
        ' Load image resource
        Set G_oDDResourceFlame = LoadBitmapIntoDXS(App.Path + "\flame.bmp")
        
        ' Flame #1 ...
        
            ' Create frame to contain decal and lighting, position frame
            G_oD3DInstance.CreateFrame G_oD3DMasterFrame, G_oD3DFrameFlame1
            G_oD3DFrameFlame1.SetPosition Nothing, 12.45, 9.5, 1.5
            
            ' Create surface to contain current animation
            Set G_oDDSurfaceFlame1 = MakeDXSurface(32, 32)
            
            ' Create decal texture from surface
            G_oD3DInstance.CreateTextureFromSurface G_oDDSurfaceFlame1, G_oD3DTextureFlame1
            
            ' Set decal texture properties
            With G_oD3DTextureFlame1
                .SetDecalOrigin 16, 8
                .SetDecalScale 1
                .SetDecalSize 0.5, 0.75
                .SetDecalTransparency 1
                .SetDecalTransparentColor 0
            End With
            
            ' Add decal to frame
            G_oD3DFrameFlame1.AddVisual G_oD3DTextureFlame1
            
        ' Flame #2 ...
        
            ' Create frame to contain decal and lighting, position frame
            G_oD3DInstance.CreateFrame G_oD3DMasterFrame, G_oD3DFrameFlame2
            G_oD3DFrameFlame2.SetPosition Nothing, 9.5, 9.5, 1.5

            ' Create surface to contain current animation
            Set G_oDDSurfaceFlame2 = MakeDXSurface(32, 32)
            
            ' Create decal texture from surface
            G_oD3DInstance.CreateTextureFromSurface G_oDDSurfaceFlame2, G_oD3DTextureFlame2
            
            ' Set decal texture properties
            With G_oD3DTextureFlame2
                .SetDecalOrigin 16, 8
                .SetDecalScale 1
                .SetDecalSize 0.5, 0.75
                .SetDecalTransparency 1
                .SetDecalTransparentColor 0
            End With
            
            ' Add decal to frame
            G_oD3DFrameFlame2.AddVisual G_oD3DTextureFlame2
        
    ' Create lighting for flames ...
        
        ' Lighting #1 ...
        
            ' Create light and set its propertys
            G_oD3DInstance.CreateLightRGB D3DRMLIGHT_POINT, 0.4, 0.3, 0.6, G_oD3DLightFlame1
            G_oD3DLightFlame1.SetConstantAttenuation 1
            
            ' Add light to frame
            G_oD3DFrameFlame1.AddLight G_oD3DLightFlame1
        
        ' Lighting #2 ...
        
            ' Create light and set its propertys
            G_oD3DInstance.CreateLightRGB D3DRMLIGHT_POINT, 0.4, 0.3, 0.6, G_oD3DLightFlame2
            G_oD3DLightFlame2.SetConstantAttenuation 1
            
            ' Add light to frame
            G_oD3DFrameFlame2.AddLight G_oD3DLightFlame2
    
    ' Error handling ...
    
        Exit Sub

E_CreateFlames:

        AppError Err.Number, Err.Description, "CreateFlames"
    
End Sub
    
Private Sub CreateMirror()


    ' Enable error handling...
        On Error GoTo E_CreateMirror

    ' Setup local variables ...
        Dim L_oD3DFace As IDirect3DRMFace                   ' Face to be added to meshbuilder
        Dim L_oD3DMeshbuilder As IDirect3DRMMeshBuilder2    ' Meshbuilder to hold created face
        
    ' Create mirror surface, texture and material ...
        
        ' Create surface to hold mirror
        Set G_oDDSurfaceMirror = MakeDXSurface(64, 64, True)
        
        ' Create mirror texture from surface
        G_oD3DInstance.CreateTextureFromSurface G_oDDSurfaceMirror, G_oD3DTextureMirror
        
        ' Create mirror material
        G_oD3DInstance.CreateMaterial 1, G_oD3DMaterialMirror
        G_oD3DMaterialMirror.SetEmissive 0.4, 0.4, 0.4
        
    ' Initialize mirror display system
    
        ' Create a D3DRM device from the mirror surface
        G_oD3DInstance.CreateDeviceFromSurface G_dD3DDriver.GUID, G_oDDInstance, G_oDDSurfaceMirror, G_oD3DDeviceMirror
    
        ' Check device existance, terminate if missing
        If G_oD3DDeviceMirror Is Nothing Then
           AppError 0, "Could not create D3DRM device", "CreateMirror"
           Exit Sub
        End If
    
        ' Set D3DRM device quality
        G_oD3DDeviceMirror.SetQuality D3DRMLIGHT_ON Or D3DRMFILL_SOLID Or D3DRMSHADE_GOURAUD
        
        ' Create the camera frame containing the mirror camera
        G_oD3DInstance.CreateFrame G_oD3DMasterFrame, G_oD3DFrameMirror
        
        ' Create a D3D viewport from the device, using the camera frame for output
        G_oD3DInstance.CreateViewport G_oD3DDeviceMirror, G_oD3DFrameMirror, 0, 0, 64, 64, G_oD3DViewportMirror
            
        ' Check viewport existance, terminate if missing
        If G_oD3DViewportMirror Is Nothing Then
           AppError 0, "Could not create D3DRM viewport", "CreateMirror"
           Exit Sub
        End If
        
        ' Set the projection model and properties for the viewport
        With G_oD3DViewportMirror
            .SetProjection D3DRMPROJECT_PERSPECTIVE
            .SetBack 10
            .SetFront 1
        End With
        
        ' Set initial mirror camera orientation and position
        G_oD3DFrameMirror.SetPosition Nothing, 11, 13, 2
        G_oD3DFrameMirror.SetOrientation Nothing, 0, -1, -0.2, 0, 0, 1
        
    ' Create faces holding mirror ...
    
        ' Initialize meshbuilder
        G_oD3DInstance.CreateMeshBuilder L_oD3DMeshbuilder

        ' Create mirror surrounding (forward looking) ...
        
            ' Create face
            G_oD3DInstance.CreateFace L_oD3DFace
            With L_oD3DFace
                .AddVertex 9.9, 12.5, 1
                .AddVertex 12.1, 12.5, 1
                .AddVertex 12.1, 12.5, 3.2
                .AddVertex 9.9, 12.5, 3.2
                .SetTextureCoordinates 0, 11, 2
                .SetTextureCoordinates 1, 10, 2
                .SetTextureCoordinates 2, 10, 1
                .SetTextureCoordinates 3, 11, 1
            End With
                    
            ' Set mirror face texture and material
            L_oD3DFace.SetColorRGB 0.1, 0.1, 0.1
            
            ' Add face data to meshbuilder
            L_oD3DMeshbuilder.AddFace L_oD3DFace
            
            ' Release face
            Set L_oD3DFace = Nothing
        
        ' Create mirror surrounding (backward looking) ...
        
            ' Create face
            G_oD3DInstance.CreateFace L_oD3DFace
            With L_oD3DFace
                .AddVertex 9.9, 12.5, 3.2
                .AddVertex 12.1, 12.5, 3.2
                .AddVertex 12.1, 12.5, 1
                .AddVertex 9.9, 12.5, 1
            End With
                    
            ' Set mirror face texture and material
            L_oD3DFace.SetColorRGB 0.1, 0.1, 0.1
            
            ' Add face data to meshbuilder
            L_oD3DMeshbuilder.AddFace L_oD3DFace
            
            ' Release face
            Set L_oD3DFace = Nothing
        
        ' Create mirror surface ...
        
            ' Create face holding mirror
            G_oD3DInstance.CreateFace L_oD3DFace
            With L_oD3DFace
                .AddVertex 10, 12.45, 1.1
                .AddVertex 12, 12.45, 1.1
                .AddVertex 12, 12.45, 3.1
                .AddVertex 10, 12.45, 3.1
                .SetTextureCoordinates 0, 10, 2
                .SetTextureCoordinates 1, 11, 2
                .SetTextureCoordinates 2, 11, 1
                .SetTextureCoordinates 3, 10, 1
            End With
                    
            ' Set mirror face texture and material
            L_oD3DFace.SetTexture G_oD3DTextureMirror
            L_oD3DFace.SetMaterial G_oD3DMaterialMirror
            
            ' Add face data to meshbuilder
            L_oD3DMeshbuilder.AddFace L_oD3DFace
                
            ' Reset face
            Set L_oD3DFace = Nothing
        
        ' Generate lighting normals
        L_oD3DMeshbuilder.GenerateNormals
        
        ' Add created mesh to frame
        G_oD3DMasterFrame.AddVisual L_oD3DMeshbuilder
     
        ' Clean up all DirectX objects
        Set L_oD3DMeshbuilder = Nothing
    
    ' Error handling ...
    
        Exit Sub

E_CreateMirror:
    
        AppError Err.Number, Err.Description, "CreateMirror"
    
End Sub


Private Sub UpdateChars()

    ' Enable error handling ...
        
        On Error GoTo E_UpdateChars
    
    ' Setup local variables ...
        
        Dim L_dRenderArea As RECT           ' Variable holding blitting area
        Dim L_nRunChar As Integer           ' Variable to run through all characters to be displayed
        Dim L_nRunRow As Integer            ' Variable to run through rows within a character
        Dim L_nRunCol As Integer            ' Variable to run through columns within a character
        Dim L_dDDBLTFX As DDBLTFX           ' FX descriptor for blitting
        Dim L_nColorFactor As Long          ' Color factor for blitting characters
        
    ' Update scrolling characters ...
        
        ' Scroll text
        G_nCharScrollPos = G_nCharScrollPos + 5
        
        ' If offset of one character reached
        If G_nCharScrollPos > 36 Then
            
            ' Reset offset
            G_nCharScrollPos = 0
            
            ' Scroll text
            G_sDisplayText = Right(G_sDisplayText, Len(G_sDisplayText) - 1) + Left(G_sDisplayText, 1)
            
            ' Rebuild surface holding characters ...
            
                ' Clear surface
                With L_dDDBLTFX
                    .dwFillColor = 0
                    .dwSize = Len(L_dDDBLTFX)
                End With
                With L_dRenderArea
                    .Top = 0
                    .Left = 0
                    .Bottom = 47
                    .Right = 520
                End With
                G_oDDSurfaceChars.Blt L_dRenderArea, ByVal Nothing, ByVal 0&, DDBLT_COLORFILL, L_dDDBLTFX
                
                ' Draw characters...
                
                ' Run through all characters, and through rows and columns of the characters
                For L_nRunChar = 0 To 12
                    For L_nRunRow = 0 To 6
                        For L_nRunCol = 0 To 4
                            
                            ' If the font data set tells that a pixel is to be drawn
                            If G_bFontData(Asc(Mid(G_sDisplayText, L_nRunChar + 1, 1)), L_nRunRow * 5 + L_nRunCol) Then
                                
                                ' Set fillcolor according to row position
                                L_nColorFactor = (31 - 5 * Abs(3 - L_nRunRow))
                                L_dDDBLTFX.dwFillColor = L_nColorFactor * 1024 + L_nColorFactor * 32 + L_nColorFactor
                                
                                ' Set render area according to current position
                                With L_dRenderArea
                                    .Top = L_nRunRow * 7
                                    .Bottom = .Top + 5
                                    .Left = L_nRunChar * 40 + L_nRunCol * 7
                                    .Right = .Left + 5
                                End With
                                
                                ' Blit to character surface
                                G_oDDSurfaceChars.Blt L_dRenderArea, ByVal Nothing, ByVal 0&, DDBLT_COLORFILL, L_dDDBLTFX
                                
                            End If
                            
                        Next
                    Next
                Next

        End If
        
        With L_dRenderArea
            .Top = 0
            .Left = G_nCharScrollPos
            .Bottom = .Top + 47
            .Right = .Left + 480
        End With
        G_oDDBackbuffer.BltFast G_dDDWindow(1).nX, G_dDDWindow(1).nY, G_oDDSurfaceChars, L_dRenderArea, DDBLTFAST_SRCCOLORKEY
        
    ' Error handler ...
        Exit Sub
    
E_UpdateChars:

    AppError Err.Number, Err.Description, "UpdateChars"

End Sub

Private Sub UpdateWindows()

    ' Enable error handling ...
        
        On Error GoTo E_UpdateWindows

    ' Setup local variables ...
        
        Dim L_nRunWindows As Integer      ' Variable to run through all windows
        
    ' Update all windows that are activated ...
    
        For L_nRunWindows = 0 To 2
            With G_dDDWindow(L_nRunWindows)
            
                ' Update window position
                .nX = .nX + .nDX
                .nY = .nY + .nDY

                ' Reflect window on edges
                With G_dDDWindow(L_nRunWindows)
                    If .nY > 470 - (.dRenderArea.Bottom - .dRenderArea.Top) Then .nDY = -.nDY
                    If .nY < 10 Then .nDY = -.nDY
                    If .nX > 630 - (.dRenderArea.Right - .dRenderArea.Left) Then .nDX = -.nDX
                    If .nX < 10 Then .nDX = -.nDX
                 End With
                                     
                ' Redraw window contents
                If Not (G_dDDWindow(L_nRunWindows).oDDSurface Is Nothing) Then
                    G_oDDBackbuffer.BltFast .nX, .nY, .oDDSurface, .dRenderArea, DDBLTFAST_SRCCOLORKEY
                End If
                
            End With
        Next
        
    ' Error handler ...
        Exit Sub
    
E_UpdateWindows:

        AppError Err.Number, Err.Description, "UpdateWindows"
                
End Sub

    

Private Sub UpdateGround()

    ' Enable error handling ...
        
        On Error GoTo E_UpdateGround
    
    ' Setup local variables ...
        
        Dim L_nSeperator As Integer           ' Holds current scroll seperator position
        Dim L_dRenderArea As RECT             ' Variable holding blitting area
        
    ' Update scrolling lava texture
    
        ' Render current phase to surface the texture is attached to
        L_nSeperator = 32 - (G_nFrameCount Mod 32)
        With L_dRenderArea
            .Top = 0
            .Left = 0
            .Bottom = L_nSeperator
            .Right = 32
        End With
        If L_nSeperator > 1 Then G_oDDSurfaceLava.BltFast 0, 32 - L_nSeperator, G_oDDResourceLava, L_dRenderArea, DDBLTFAST_NOCOLORKEY
        With L_dRenderArea
            .Top = L_nSeperator + 1
            .Left = 0
            .Bottom = 32
            .Right = 32
        End With
        If L_nSeperator < 31 Then G_oDDSurfaceLava.BltFast 0, 0, G_oDDResourceLava, L_dRenderArea, DDBLTFAST_NOCOLORKEY
            
        ' Inform D3DRM that the surface the texture is attached to has changed
        G_oD3DTextureLava.Changed 1, 0

    ' Error handler ...
        Exit Sub
    
E_UpdateGround:

        AppError Err.Number, Err.Description, "UpdateGround"

End Sub

Private Sub UpdateFlames()

    ' Enable error handling ...
        
        On Error GoTo E_UpdateFlames
    
    ' Setup local variables ...
        Dim L_dRenderArea As RECT       ' Variable holding blitting area

        ' Update flame lighting #1
        G_oD3DLightFlame1.SetConstantAttenuation 0.5 + Rnd

        ' Update texture for flame #1
        With L_dRenderArea
            .Left = ((G_nFrameCount / 2) Mod 4) * 32
            .Top = Int(((G_nFrameCount / 2) Mod 16) / 4) * 32
            .Bottom = .Top + 32
            .Right = .Left + 32
        End With
        G_oDDSurfaceFlame1.BltFast 0, 0, G_oDDResourceFlame, L_dRenderArea, DDBLTFAST_NOCOLORKEY
    
        ' Inform D3DRM that the surface the texture is attached to has changed
        G_oD3DTextureFlame1.Changed 1, 0
        
        ' Update flame lighting #2
        G_oD3DLightFlame2.SetConstantAttenuation 0.5 + Rnd
        
        ' Update texture for flame #2
        With L_dRenderArea
            .Left = ((G_nFrameCount / 2) Mod 4) * 32
            .Top = Int(((G_nFrameCount / 2) Mod 16) / 4) * 32
            .Bottom = .Top + 32
            .Right = .Left + 32
        End With
        G_oDDSurfaceFlame2.BltFast 0, 0, G_oDDResourceFlame, L_dRenderArea, DDBLTFAST_NOCOLORKEY
    
        ' Inform D3DRM that the surface the texture is attached to has changed
        G_oD3DTextureFlame2.Changed 1, 0

    ' Error handler ...
        Exit Sub
    
E_UpdateFlames:

        AppError Err.Number, Err.Description, "UpdateFlames"

End Sub

Private Sub UpdateScene()

    ' Enable error handling ...
        
        On Error GoTo E_UpdateScene
    
    ' Setup local variables ...

        ' Set the new camera position from the lookup table, loop position within lookup table
        G_nCamPosCurrent = G_nCamPosCurrent + 1
        If G_nCamPosCurrent > 179 Then G_nCamPosCurrent = 0
        
        ' Set camera position and orientation to new values
        With G_dCamPosLookup(G_nCamPosCurrent)
            G_oD3DCameraFrame.SetPosition Nothing, .X, .Y, .z
            G_oD3DCameraFrame.SetOrientation Nothing, 11 - .X, 10.5 - .Y, -4, 0, 0, 1
        End With
    
        ' Update D3DRM model
        G_oD3DInstance.Tick 1
        
    ' Error handler ...
        Exit Sub
    
E_UpdateScene:

        AppError Err.Number, Err.Description, "UpdateScene"

End Sub

Private Sub UpdateMirror()

    ' Enable error handling ...
        
        On Error GoTo E_UpdateMirror
    
    ' Update mirror animation ...
        
        ' Inform the renderer that the texture has changed
        G_oD3DTextureMirror.Changed 1, 0
        
    ' Error handler ...
        Exit Sub
    
E_UpdateMirror:

        AppError Err.Number, Err.Description, "UpdateMirror"

End Sub


Private Sub UpdateBack()

    ' Enable error handling...
        On Error GoTo E_UpdateBack

    ' Setup local variables ...
            
        Dim L_dDDSD As DDSURFACEDESC    ' Description of surface to be obtained by lock
        Dim L_nSurfacePointer As Long   ' Pointer to the surface
        Dim L_nRun As Integer           ' Variable to run through arrays
        Dim L_dRenderArea As RECT       ' Area to render explosions from
        
    ' Update background stars ...
    
        ' Prepare structure to obtain lock
        L_dDDSD.dwSize = Len(L_dDDSD)
        L_dDDSD.dwFlags = DDSD_LPSURFACE

        ' Obtain lock to surface
        G_oDDBackbuffer.Lock ByVal 0&, L_dDDSD, DDLOCK_SURFACEMEMORYPTR, ByVal 0&
            
        ' Get pointer to surface memory
        L_nSurfacePointer = L_dDDSD.lpSurface
        
        ' Calculate and draw stars ...
        For L_nRun = 0 To 1999
            With G_dStar(L_nRun)
                
                .nX = .nX - .nSpeed
                If .nX < 5 Then .nX = 635
                
                CopyMemory ByVal (L_nSurfacePointer + .nX * 2 + .nY * 1280), ByVal VarPtr(.nColor), 2
                
            End With
        Next
        
        ' Release lock to surface
        G_oDDBackbuffer.Unlock ByVal 0&
        
    ' Update background explosions ...
    
        For L_nRun = 0 To 14
            With G_dExplo(L_nRun)
                    
                ' Calculate new position and phase
                .nPhase = .nPhase + 1
                .nX = .nX - 3
                
                ' Spawn new explosion if old one off screen or finished ...
                If .nPhase > 15 Or .nX < 3 Then
                    
                    ' Set visual data
                    .nPhase = 0
                    .nX = Int(Rnd * 600) + 5
                    .nY = Int(Rnd * 470) + 5
                    
                    ' Set sound effect for first 8 explosions ...
                    If L_nRun < 8 Then
                        
                        ' Stop if still playing
                        .oDSBuffer.Stop
                        
                        ' Alter frequency randomly, resulting in various pitch
                        .oDSBuffer.SetFrequency ByVal Int(Rnd * 35000) + 15000
                        
                        ' Set pan to fit position , resulting in stereo effect
                        .oDSBuffer.SetPan (320 - .nX) * 10
                        
                        ' Alter volume randomly, resulting in distinguishable explosions
                        .oDSBuffer.SetVolume ByVal (-300 - Int(Rnd * 400))
                        
                        ' Restart sound
                        .oDSBuffer.Play ByVal 0&, ByVal 0&, ByVal 0&
                        
                    End If
                    
                End If
                
                ' Calculate render area, clip at edges
                L_dRenderArea.Top = 0
                L_dRenderArea.Left = .nPhase * 32 + IIf(.nX < 32, 32 - .nX, 0)
                L_dRenderArea.Right = L_dRenderArea.Left + IIf(.nX < 32, .nX, 32)
                L_dRenderArea.Bottom = L_dRenderArea.Top + IIf(480 - .nY < 32, 480 - .nY, 32)
                
                ' Blit explosion to backbuffer
                G_oDDBackbuffer.BltFast .nX, .nY, G_oDDSurfaceExplo, L_dRenderArea, DDBLTFAST_SRCCOLORKEY
                
            End With
        Next
        
    ' Error handling ...
    
        Exit Sub

E_UpdateBack:

        AppError Err.Number, Err.Description, "UpdateBack"
        Exit Sub

End Sub

