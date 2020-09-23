VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Region Generator"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6600
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox txtSource 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   7455
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   1560
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open Picture"
      Filter          =   "All Image files|*.bmp;*.dib;*.jpg;*.gif;*.wmf;*.emf"
      Flags           =   2627588
   End
   Begin VB.PictureBox pctPicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      Height          =   1095
      Left            =   5040
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   155
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label txtStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6180
      Width           =   6255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Easy Region Generator Demo
' --------------------------
' Don't let the length of the code scare you away.
' Everything happens in CreateRegion and cmdTest_Click.
' The rest is just extra stuff to help you reuse the region
' in other projects.

' CreateRegion - Creates the Region
' ------------------
' This function has been optimised after the suggestion
' of Will Barden posted on PSC on 7/24/2003 2:55:56 PM
' Will suggested that instead of creating 1x1 regions
' whenever a transparent pixel is found, create one region
' for successive transparent pixels when a non-transparent
' pixel is found. For example if 5 transparent pixels are
' found before a non-transparent pixel, create a region of
' 1x5 pixels instead of 5 regions of 1x1. However, the speed
' gain always depends on the size and complexity of the picture.
' To enable optimisation set the following
' "conditional compiler constant" to true

#Const OPTIMISED = False

Option Explicit

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long


Private Const RGN_XOR = 3

Dim b()    As Byte  ' Region data bytes
Dim nBytes As Long  ' Number of region data bytes

' When the user clicks open
Private Sub cmdOpen_Click()
    On Local Error GoTo errOpen
    
    ' Show the open dialog
    cdlg.ShowOpen
    
    ' Change the mouse pointer to busy
    MousePointer = vbHourglass
    
    ' Try to open the picture
    pctPicture.Picture = LoadPicture(cdlg.Filename)
    
    ' Create the Region
    CreateRegion vbMagenta
    
    ' Generate the source code
    GenerateSourceCode
    
    ' Create Binary file
    CreateBinaryFile

errOpen:
    ' Change the mouse pointer back to normal
    MousePointer = vbDefault
    
    ' Enable the Test button depending on the error condition
    ' And notify the user
    If Err.Number = 0 Then
        cmdTest.Enabled = True
        SetStatus "Done"
    Else
        cmdTest.Enabled = False
        SetStatus "Nothing Happened!!!"
    End If
    Err.Clear
End Sub

' Creates the Region
Private Sub CreateRegion(TransparentColor As Long)
    Dim hRgn     As Long    ' Our regions
    Dim X        As Long    ' X coordinate of the pixel checked
    Dim Y        As Long    ' Y coordinate of the pixel checked
    
    ' Change the status message
    SetStatus "Creating Region..."
    
    ' Create an initial rectangular region
    hRgn = CreateRectRgn(0, 0, pctPicture.ScaleWidth, pctPicture.ScaleHeight)
    
    ' Initialise the Progress Bar
    pb.Min = 0
    pb.Max = pctPicture.ScaleWidth + 1
    pb.Value = 0
    
' Compiling the code with OPTIMISED set to true will
' leave the #else part out and if OPTIMISED is set to
' another value it will leave the #if part out
' When run from the VB IDE, it will only execute the part
' that evaluates to true. If you need more help with this
' feature feel free to ask me :)
#If OPTIMISED = True Then

    Dim XCount As Long  ' Number of successive transparent pixels on X-axis
    Dim XStart As Long  ' Start of successive transparent pixels found in a line
    
    ' Remove any pixels that should not be in the region
    For Y = 0 To pctPicture.ScaleHeight
        For X = 0 To pctPicture.ScaleWidth
            If pctPicture.Point(X, Y) = TransparentColor Then
                XCount = XCount + 1
            Else
                ExcludePixelsFromRegion hRgn, XStart, XCount, X, Y
            End If
        Next X
        
        ' When the end of the line has been reached and there are transparent
        ' pixels to be excluded
        If XCount <> 0 Then
            ExcludePixelsFromRegion hRgn, XStart, XCount, X, Y
        End If
                
        pb.Value = pb.Value + 1
    Next Y

#Else

    Dim hRgnTemp As Long    ' Temporary region to exclude

    ' Remove any pixels that should not be in the region
    For X = 0 To pctPicture.ScaleWidth
        For Y = 0 To pctPicture.ScaleHeight
            If pctPicture.Point(X, Y) = TransparentColor Then
                hRgnTemp = CreateRectRgn(X, Y, X + 1, Y + 1)
                CombineRgn hRgn, hRgn, hRgnTemp, RGN_XOR
                DeleteObject hRgnTemp
            End If
        Next Y
        pb.Value = pb.Value + 1
    Next X

#End If
    
    ' Get the region data so that we can use them again
    nBytes = GetRegionData(hRgn, 0, ByVal 0&)
    ReDim b(nBytes)
    GetRegionData hRgn, nBytes, b(0)
    
    ' Delete the region handle
    DeleteObject hRgn
End Sub

' Excludes a line of XCount pixels from the final region
Private Sub ExcludePixelsFromRegion(hRgn As Long, XStart As Long, XCount As Long, X As Long, Y As Long)
    Dim hRgnTemp As Long    ' Temporary region to exclude
    
    hRgnTemp = CreateRectRgn(XStart, Y, XStart + XCount, Y + 1)
    CombineRgn hRgn, hRgn, hRgnTemp, RGN_XOR
    DeleteObject hRgnTemp
    XCount = 0
    XStart = (X + 1) Mod (pctPicture.ScaleWidth + 1)
End Sub

' Generates the source code needed to create this region
Private Sub GenerateSourceCode()
    Dim i        As Long    ' Used with loops
    Dim Filename As String  ' File with the source code
    Dim buf      As String  ' Temporary buffer to read from the file
    
    ' Change the status message
    SetStatus "Generating Source Code..."
    
    ' Open a file to write the Source code
    Filename = GetFileName("Source.txt")
    
    Open Filename For Output As 1
    
    ' WINAPI Types and Declarations
    AddToSource "Private Type RECT", False
    AddToSource "Left   As Long"
    AddToSource "Top    As Long"
    AddToSource "Right  As Long"
    AddToSource "Bottom As Long"
    AddToSource "End Type", False
    AddToSource "", False
    AddToSource "Private Declare Function DeleteObject Lib ""gdi32"" (ByVal hObject As Long) As Long", False
    AddToSource "Private Declare Function ExtCreateRegion Lib ""gdi32"" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long", False
    AddToSource "Private Declare Function GetRgnBox Lib ""gdi32"" (ByVal hRgn As Long, lpRect As RECT) As Long", False
    AddToSource "Private Declare Function SetWindowRgn Lib ""user32"" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long", False
    AddToSource "", False
    
    ' The function
    AddToSource "Private Sub CreateRegion(frm as Form)", False
    AddToSource "Dim b(" & nBytes & ") as Byte"
    AddToSource "Dim hRgn as Long"
    AddToSource "Dim rc   as RECT"
    AddToSource "", False
    
    ' Initialise the Progress bar
    pb.Min = 0
    pb.Max = nBytes + 1
    pb.Value = 0
    
    For i = 0 To nBytes
        If b(i) <> 0 Then AddToSource "b(" & i & ") = " & b(i)
        pb.Value = pb.Value + 1
    Next i
            
    AddToSource "hRgn = ExtCreateRegion(ByVal 0&, " & nBytes & ", b(0))"
    AddToSource "GetRgnBox hRgn, rc" & vbNewLine
    AddToSource "SetWindowRgn frm.hWnd, hRgn, True"
    AddToSource "DeleteObject hRgn"
    AddToSource "frm.Width = rc.Right * Screen.TwipsPerPixelX"
    AddToSource "frm.Height = rc.Bottom * Screen.TwipsPerPixelY"
    AddToSource "End Sub", False
    Close 1
    
    ' Open the file to read and display the contents (source code)
    Open Filename For Binary As 1
        buf = Space(LOF(1))
        Get 1, , buf
        txtSource = buf
    Close 1
    
    If Len(txtSource) < Len(buf) Then
        MsgBox "Source code has been truncated" & vbNewLine & _
        "You might not be able to use this source code because" & vbNewLine & _
        "it is very large. Alternatively you can use the generated" & vbNewLine & _
        "binary file. See example code for details.", vbInformation, "Truncated Source"
    End If
End Sub

' Creates a Binary file with the region information
Private Sub CreateBinaryFile()
    Dim Filename As String
    Dim i        As Long
    
    ' Reset the progress bar
    pb.Min = 0
    pb.Max = nBytes + 1
    pb.Value = 0
    
    ' Set the status
    SetStatus "Creating Binary File"
    
    Filename = GetFileName("Binary.bin")
    
    Open Filename For Binary As 1
    
    For i = 0 To nBytes
        Put 1, , b(i)
        pb.Value = pb.Value + 1
    Next i
    
    Close 1
End Sub

' When the user clicks on the test button
Private Sub cmdTest_Click()
    Dim hRgn As Long    ' Our region
    Dim rc   As RECT    ' The bounding rectangle of the region
    
    ' Set the window picture - this causes the window to load as well
    frmTest.pctTest.Picture = pctPicture
    
    ' Create, Measure, Set and Delete the region
    hRgn = ExtCreateRegion(ByVal 0&, nBytes, b(0))
    GetRgnBox hRgn, rc
    SetWindowRgn frmTest.hWnd, hRgn, True
    DeleteObject hRgn

    ' Resize the window and show it
    frmTest.Width = rc.Right * Screen.TwipsPerPixelX
    frmTest.Height = rc.Bottom * Screen.TwipsPerPixelY
    Me.Hide
    frmTest.Show vbModal, Me
    Me.Show
End Sub

' Adds a line to the source code text box
Private Sub AddToSource(Text As String, Optional bTab As Boolean = True)
    If bTab Then
        Print #1, vbTab & Text
    Else
        Print #1, Text
    End If
End Sub

' Changes the Status message
Private Sub SetStatus(Text As String)
    txtStatus = "Status : " & Text
    Refresh
End Sub

' Returns a full path for the filename and deletes any
' files found with the same name
Private Function GetFileName(File As String, Optional DeletePrevious As Boolean = True) As String
    Dim Filename As String
    
    Filename = App.Path
    If Right(Filename, 1) = "\" Then
        Filename = Filename & File
    Else
        Filename = Filename & "\" & File
    End If
    
    If DeletePrevious Then
        If Dir(Filename, vbNormal) <> "" Then Kill Filename
    End If
    
    GetFileName = Filename
End Function

