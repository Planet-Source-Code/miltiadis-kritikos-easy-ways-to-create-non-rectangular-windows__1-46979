VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Binary File Region Example"
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   0
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   4500
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

' Add this code to move the form with the mouse
Private Sub pctPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub


' Code to exit on ESC keypress
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim b()      As Byte
    Dim nBytes   As Long

    ReadData b, nBytes
    CreateRegion b, nBytes
End Sub

' Reads the data from the file
Private Sub ReadData(b() As Byte, nBytes As Long)
    Dim FileName As String
    Dim i        As Long
    
    ' Get and open the filename with the region data
    FileName = GetFileName("Binary.bin")
    Open FileName For Binary As 1
    
    ' Get the number of bytes in this file
    nBytes = LOF(1)
    
    ' Resize the binary array and read the data, close the file in the end
    ReDim b(nBytes)
    For i = 0 To nBytes
        Get 1, , b(i)
    Next i
    Close 1
End Sub

' Creates the region
Private Sub CreateRegion(b() As Byte, nBytes As Long)
    Dim hRgn As Long    ' Our region
    Dim rc   As RECT    ' The bounding rectangle of the region
    
    ' Create, Measure, Set and Delete the region
    hRgn = ExtCreateRegion(ByVal 0&, nBytes, b(0))
    GetRgnBox hRgn, rc
    SetWindowRgn hWnd, hRgn, True
    DeleteObject hRgn

    ' Resize the window and show it
    Width = rc.Right * Screen.TwipsPerPixelX
    Height = rc.Bottom * Screen.TwipsPerPixelY
End Sub

' Returns a full path for the filename and deletes any
' files found with the same name
Private Function GetFileName(File As String) As String
    Dim FileName As String
    
    FileName = App.Path
    If Right(FileName, 1) = "\" Then
        FileName = FileName & File
    Else
        FileName = FileName & "\" & File
    End If
        
    GetFileName = FileName
End Function
