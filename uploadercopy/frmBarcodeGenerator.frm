VERSION 5.00
Begin VB.Form frmBarcodeViewer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EAN Generator"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3600
   ForeColor       =   &H8000000C&
   Icon            =   "frmBarcodeGenerator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBarcode 
      Height          =   315
      Left            =   120
      MaxLength       =   13
      TabIndex        =   2
      Tag             =   "Enter 7+ digits"
      ToolTipText     =   "Enter 7+ digits"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdEANCreate 
      Caption         =   "&Generate"
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Click here to generate a barcode"
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picEan 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      ScaleHeight     =   80
      ScaleMode       =   0  'User
      ScaleWidth      =   113
      TabIndex        =   0
      Top             =   660
      Width           =   3375
   End
   Begin VB.Menu zMenu1 
      Caption         =   "&Ean"
      Begin VB.Menu zEan 
         Caption         =   "&Generate"
         Index           =   0
         Shortcut        =   ^G
      End
      Begin VB.Menu zEan 
         Caption         =   "&Save Ean"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu zEan 
         Caption         =   "&Print"
         Index           =   2
         Shortcut        =   ^P
      End
      Begin VB.Menu zEan 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu zEan 
         Caption         =   "&Quit"
         Index           =   4
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmBarcodeViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Original spiel by original author:
'==================================
'Well, not many functions implemented here. I tried to make it as simple
'as possible even if this app loses any practical usage due it's simplicity,
'exept for the check number calculation and that's just what I needed.
'Be patient with my english,- I am not from this world.
'Any questions? Then mail me: davidsmejkal@hellada.cz

'Updated spiel by some guy ;) :
'==============================
'I've tried to make this as close to this VB naming convention file that i found
' as possible. Can somebody tell me why using short names is bad? Is the file
'size affected by the size of variable name that you use? Any comments welcome.
'
'New features:
'Includes a different checkdigit function than the original.
'Has been updated to include support for EAN 8 barcodes. SO FAR all generated
'barcodes have been good, any problems please email!
'Tried and tested at my local Somerfields :D
'Please email the updater: cmeister2@hotmail.com
'Or: cmeister2@btinternet.com (more likely to be replied to, since hotmail's
'filter goes mad on real emails to me :D


Dim m_sBarcode As String, m_lBarcodeLength As Long

Private Sub cmdEANCreate_Click()
On Error GoTo errHandler                            'Error Handling function

Dim bytBarcodeType As Byte, sTemp As String         'Initiate variables
With txtBarcode
Select Case Len(.Text)
    Case 0 To 6:
        Alert "Enter 7+ numbers into the text box": Exit Sub    '6 or less numbers entered
    Case 7 To 11:
        bytBarcodeType = 7                                      'EAN 8 barcode
        m_lBarcodeLength = 8
    Case 12 To 20:
        bytBarcodeType = 12                                     'EAN 13 barcode
        m_lBarcodeLength = 13
End Select

m_sBarcode = MakeBarcode(Left(.Text, bytBarcodeType))           'Puts correct checkdigit on barcode root.
.Text = m_sBarcode                                              'Full EAN code
DrawEan                                                         'Draw the barcode!

End With
Exit Sub

errHandler:
Select Case Err.Number
    Case 13: Alert "Enter only numbers into text box!"   'In case someone puts other characters then numbers into textbox
    Case Else: Alert "Error occurred: " & Err.Description   'Any other error, die nicely
End Select
End Sub

Private Sub Form_Load()
Init                            'Initializes Mdl array - this holds the lines info!
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmBarcodeViewer = Nothing        'Important!!! = remove form
End Sub

Private Sub DrawEan()

Dim bytCentreDigit As Byte, lPositionX As Long, i As Integer, j As Integer
Dim lCurrNumber As Long, lFirstNumber As Long, iModule As Integer

bytCentreDigit = IIf(m_lBarcodeLength = 8, 5, 8)     'Where to put the middle bars? EAN8: 5 digit, EAN13: 8th digit (just before each)
With picEan
    .Cls                                             'Clear
    .BackColor = vbWhite                             'Set colour
    .FontSize = 12                                   'Set font size
    .DrawWidth = 2                                   'Set draw width
lPositionX = 11                   'X position (11 =must be 11 modules [1 module = usually 0.33 millimeters, in my case picEan.ScaleWidth <bar width> / 113] 11 on left side, 7 on right side

For i = 1 To m_lBarcodeLength     '8 or 13 digit code
    lCurrNumber = CLng(Mid(m_sBarcode, i, 1)) 'Current n°
    If i = 1 Then
        GuardBar lPositionX         'Draw double lines at current X position
        lFirstNumber = lCurrNumber  'This
        .CurrentX = 2
        .CurrentY = 66
        picEan.Print IIf(m_lBarcodeLength = 8, "<", lFirstNumber) 'If EAN8, draw "<", else draw number
    End If
    If i <> 1 Or m_lBarcodeLength = 8 Then
        If i < bytCentreDigit Then                'On the left side, there are modules 1 or 2 (A, B) depending on the 1st digit = [Mdl(0 - 9, 0 or 1)]...
        Select Case m_lBarcodeLength
            Case 8: iModule = 0                   'For EAN 8, always use module 0 (if doesnt work, see start for email addy. Please inform!
            Case 13: iModule = MidInt(MdlLeft(lFirstNumber), i - 1)
        End Select
        Else: iModule = 2                     '...on the right side always module 2 (C) = [Mdl(0 - 9, 2)]
        End If
        If i = bytCentreDigit Then                       'Draw the centre pattern
            lPositionX = lPositionX + 2
            GuardBar lPositionX
            lPositionX = lPositionX + 1
        End If
        For j = 1 To 7                  '7 modules for each n° (System of 7 black or white sprites)
            If MidInt(Mdl(iModule)(lCurrNumber), j) = 1 Then DrawLine lPositionX, 0 'Draw modules(sprites) for each n°
            lPositionX = lPositionX + 1
        Next j
        .CurrentX = lPositionX - 8
        .CurrentY = 66
        picEan.Print lCurrNumber                  'Print n°s
    End If
    Next i
    .CurrentX = lPositionX + 8
    .CurrentY = 66
    If m_lBarcodeLength = 8 Then picEan.Print ">"
    GuardBar lPositionX
End With
End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
    Case 13: cmdEANCreate_Click: Exit Sub
    Case 8, 48 To 57: Exit Sub              'Allows only numbers to be typed
    Case Else: KeyAscii = 0
End Select
End Sub

Private Sub zEan_Click(Index As Integer)
Dim sPath As String
Select Case Index
    Case 0: cmdEANCreate_Click
    Case 1
        If m_lBarcodeLength <> 0 Then            'Only if EAn is drawn
            sPath = App.Path & "\EAN-" & m_sBarcode & ".bmp"
            If Dir(sPath) <> "" Then Kill sPath     'If file exists
            SavePicture picEan.Image, sPath
            MsgBox "Ean saved as: " & Chr(34) & sPath & Chr(34)
        Else: Alert "Nothing to save!"
        End If
    Case 2
        If m_lBarcodeLength <> 0 Then PrintEan Else Alert "No bar code to print!"
    Case 4: Unload Me
End Select
End Sub

Private Sub PrintEan()
Dim i As Integer
On Error GoTo errHandler
With Printer
    .ColorMode = vbPRCMMonochrome
    .PrintQuality = -2              'Low quality
    .CurrentY = 200
    .CurrentX = 200
    .Font = "Courier New"
    .FontBold = True
    .FontSize = 10
    Printer.Print "EAN Code: " & m_sBarcode
    .FontBold = False
    Printer.PaintPicture picEan.Image, 200, 600
    Printer.Print
    .EndDoc
    MsgBox "Printing EAN Code: " & m_sBarcode & vbCrLf & "Port: " & .Port, vbInformation, App.Title
End With
    Exit Sub
errHandler:
    Printer.KillDoc
    MsgBox "Error Occurred: " & Err.Description, vbExclamation, "Error: " & Err.Number
End Sub

Private Sub GuardBar(r_lPositionX As Long)
DrawLine r_lPositionX, 6
DrawLine r_lPositionX + 2, 6
r_lPositionX = r_lPositionX + 3
End Sub

Private Sub DrawLine(r_lPositionX As Long, r_bytExtension As Byte)
picEan.Line (r_lPositionX, 5)-(r_lPositionX, 66 + r_bytExtension)
End Sub
