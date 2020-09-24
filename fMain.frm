VERSION 5.00
Begin VB.Form fMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'Kein
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   ControlBox      =   0   'False
   FillStyle       =   0  'Ausgefüllt
   ForeColor       =   &H00000000&
   Icon            =   "fMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   WindowState     =   2  'Maximiert
   Begin VB.Image IMG 
      Height          =   75
      Left            =   -30
      Top             =   -120
      Visible         =   0   'False
      Width           =   60
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()

    ' Initialize and start application ...
    
    ' Initialize application data
    Call AppInitialize
    
    ' Start main application loop
    Call AppLoop
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Finished
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

' Enable error handling...

    On Error Resume Next

' Terminate application...

    ' Cleanup application data
    Call AppTerminate
    
    ' Display logo
    Me.Hide
    Shell App.Path + "\nls.exe", vbNormalFocus
    
    ' Terminate application
    End
    
End Sub
