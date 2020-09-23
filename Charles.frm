VERSION 5.00
Begin VB.Form Charles 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "CM Screen Saver"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   144
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picWORK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   600
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picIMAG 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   144
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox picGRAD 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   144
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Timer tmrANI 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   4080
   End
   Begin VB.PictureBox picTEXT 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   144
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox picBACK 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   144
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox picFACE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   19440
      Left            =   2400
      Picture         =   "Charles.frx":0000
      ScaleHeight     =   1296
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   432
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   6480
   End
   Begin VB.Label lblFSZT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Font Size Tester"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1290
   End
End
Attribute VB_Name = "Charles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
        
    Dim RUNNING As Boolean
    
    Private Type Pstn
        LEFT As Integer
        TOP As Integer
        WIDTH As Integer
        HEIGHT As Integer
    End Type
        
    Dim TEXT As Pstn
    Dim PICT As Pstn
    Dim rctTEXT As RECT
    Dim Starting As POINTAPI

Private Sub Form_Load()

    'If this instance has been invoked in Screen Saver mode
    'then change the cursor's shape to it's "HourGlass"
    'configuration.
    If RunMODE = "S" Then Charles.MousePointer = vbHourglass
    'Call the routine which gets the Screen Saver's settings
    'from the values stored in the System Registry.
    Call GetValues

End Sub

Private Sub Form_Resize()

    'Alter the height and width of the drawing position of
    'the actual face image, to be equal to 85% of the screen's
    'total height, then centre the drawing position horizontally
    'on the screen.
    PICT.HEIGHT = 2 * Int(Charles.ScaleHeight * 0.85 / 2)
    PICT.WIDTH = PICT.HEIGHT
    PICT.LEFT = (Charles.ScaleWidth - PICT.WIDTH) / 2
    
    'Alter the width of text display position, to be equal to the
    'screen width, then alter the height to be 10% of the screen's
    'total height.
    TEXT.WIDTH = Charles.ScaleWidth
    TEXT.HEIGHT = 2 * Int(Charles.ScaleHeight * 0.1 / 2)
    
    'If the text is to appear below the face image, the top of
    'the face image is set to appear at 2½% of the screen's
    'height from the top of the screen, whilst the top of the
    'text is set to appear at the bottom of the face image.
    If CM.BelowPic Then
        PICT.TOP = 2 * Int(Charles.ScaleHeight * 0.025 / 2)
        TEXT.TOP = PICT.TOP + PICT.HEIGHT
    'If the text is to appear above the face image, the top of
    'the text is set to appear at 2½% of the screen's height
    'from the top of the screen, whilst the top of the face
    'image is set to appear at the bottom of the text.
    Else
        TEXT.TOP = 2 * Int(Charles.ScaleHeight * 0.025 / 2)
        PICT.TOP = TEXT.TOP + TEXT.HEIGHT
    End If
    
    'Make the height and width of all the controls used in the
    'drawing of the text, to be the same as our previously
    'calculated values for text height and width.
    lblFSZT.Move 0, TEXT.HEIGHT * 0, TEXT.WIDTH, TEXT.HEIGHT
    picBACK.Move 0, TEXT.HEIGHT * 1, TEXT.WIDTH, TEXT.HEIGHT
    picTEXT.Move 0, TEXT.HEIGHT * 2, TEXT.WIDTH, TEXT.HEIGHT
    picGRAD.Move 0, TEXT.HEIGHT * 3, TEXT.WIDTH, TEXT.HEIGHT
    picIMAG.Move 0, TEXT.HEIGHT * 4, TEXT.WIDTH, TEXT.HEIGHT
    
    'The Font Size is adjusted until the text it produces,
    'using the other given font characteristics, will fit
    'within the boundries of the picTEXT control.
    lblFSZT.Caption = CM.SavrText
    lblFSZT.FontSize = 144
    lblFSZT.FontName = CM.FontName
    lblFSZT.FontBold = CM.FontBold
    lblFSZT.FontItalic = CM.FontItlc
    lblFSZT.FontUnderline = CM.FontUndr
    Do:        lblFSZT.FontSize = lblFSZT.FontSize - 1
    Loop Until lblFSZT.WIDTH <= picTEXT.WIDTH _
           And lblFSZT.HEIGHT <= picTEXT.HEIGHT
           
    picTEXT.FontSize = lblFSZT.FontSize 'Sets up the font
    picTEXT.FontName = CM.FontName      'characteristics of the text
    picTEXT.FontBold = CM.FontBold      'which is ultimately to be
    picTEXT.FontItalic = CM.FontItlc    'displayed on the screen.
    picTEXT.FontUnderline = CM.FontUndr
    
    rctTEXT.LEFT = 0                    'This defines the area of the
    rctTEXT.TOP = 0                     'rectangle into which text
    rctTEXT.Right = TEXT.WIDTH          'will be written using the
    rctTEXT.Bottom = TEXT.HEIGHT        'API function "DrawText".
    
    If CM.TEXT.PictureFill = False Then
        'Fill up the control which holds the text's gradient colours.
        Call GradFill(picBACK, CM.TEXT.GradColour(1), CM.TEXT.GradColour(2), CM.TEXT.ColourStyle, CM.TEXT.ColourDrctn)
    Else
        If CM.TEXT.PictureType < 2 Then
            'Set the work Picture's size to the Screen size
            picWORK.WIDTH = Screen.WIDTH / Screen.TwipsPerPixelX
            picWORK.HEIGHT = Screen.HEIGHT / Screen.TwipsPerPixelY
            If CM.TEXT.PictureType = 0 Then
                'Paint the background into the Work picture
                RtnVal = PaintDesktop(picWORK.hDC)
            Else
                'Get the Handle for the desktop
                DeskTop_hWnd = GetDesktopWindow
                'Get the Device Context for the Desktop from it's Handle
                DeskTop_DC = GetDC(DeskTop_hWnd)
                'BitBlt the image of the desktop into the Work Picture
                RtnVal = BitBlt(picWORK.hDC, 0, 0, Charles.ScaleWidth, Charles.ScaleHeight, _
                                DeskTop_DC, 0, 0, vbSrcCopy)
            End If
            'Copy the portion of the image where the text will be written
            'from the Work Picture into the text background area.
            BitBlt picBACK.hDC, 0, 0, TEXT.WIDTH, TEXT.HEIGHT, _
                   picWORK.hDC, TEXT.LEFT, TEXT.TOP, vbSrcCopy
        Else
            Call PictFill(picBACK, picWORK, CM.TEXT.PictureAddr, CM.TEXT.PictureFit)
        End If
    End If
    
    If CM.BACK.PictureFill = False Then
        'Paint the gradient colours of the background onto the form.
        Call GradFill(Charles, CM.BACK.GradColour(1), CM.BACK.GradColour(2), CM.BACK.ColourStyle, CM.BACK.ColourDrctn)
    Else
        Select Case CM.BACK.PictureType
            Case 0
                'Paint the background into the Form
                RtnVal = PaintDesktop(Charles.hDC)
            Case 1
                'Get the Handle for the desktop
                DeskTop_hWnd = GetDesktopWindow
                'Get the Device Context for the Desktop from it's Handle
                DeskTop_DC = GetDC(DeskTop_hWnd)
                'BitBlt the image of the desktop onto the Form
                RtnVal = BitBlt(Charles.hDC, 0, 0, Charles.ScaleWidth, Charles.ScaleHeight, _
                                DeskTop_DC, 0, 0, vbSrcCopy)
            Case 2
                Call PictFill(Charles, picWORK, CM.BACK.PictureAddr, CM.BACK.PictureFit)
        End Select
    End If
    
    'Take an image of the portion of the form where the text will be blitted to.
    BitBlt picGRAD.hDC, 0, 0, TEXT.WIDTH, TEXT.HEIGHT, _
           Charles.hDC, TEXT.LEFT, TEXT.TOP, vbSrcCopy
    
    Charles.AutoRedraw = False
    
    'If this instance has been invoked in Screen Saver mode
    'then restore the cursor's shape to it's default "pointer"
    'position and then make it non-visible.
    If RunMODE = "S" Then
        Charles.MousePointer = vbDefault
        Call ShowMouse(False)
    End If
    
    'Get the starting cursor coordinates.
    RtnVal = GetCursorPos(Starting)
    
    tmrANI.Enabled = True

End Sub

'-----------------------------------------------------'
' Stop programming running when the mouse is clicked. '
'-----------------------------------------------------'
Private Sub Form_Click()

    'Function is disabled when in Preveiw mode.
    If RunMODE <> "P" Then RUNNING = False

End Sub

'-----------------------------------------------------------'
' Stop the programming running when a key has been pressed. '
'-----------------------------------------------------------'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'Function is disabled when in Preveiw mode.
    If RunMODE <> "P" Then
        'Only the ALT key is acceptable.
        If (Shift And vbAltMask) = 0 Then RUNNING = False
    End If

End Sub

'--------------------------------------------------------------'
' Stop the programme running when the cusor is moved a distace '
' of 8 or more pixels in one large movement.                   '
'--------------------------------------------------------------'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static LastX As Integer, LastY As Integer
    
    'Function is disabled when in Preveiw mode.
    If RunMODE <> "P" Then
        'If this is the first movement we set LastX and LastY
        'to their respective start coordinates.
        If LastX = 0 And LastY = 0 Then
            LastX = Starting.X
            LastY = Starting.Y
        End If
        'Use Pythagoras' theorem to determine if the cursor
        'has moved a distance of 8 or more pixels.
        If (((X - LastX) ^ 2 + (Y - LastY) ^ 2) ^ 0.5) >= 8 Then
            'If it's a large movement then stop the programme.
            RUNNING = False
        Else
            'If it's a small movement then save current coordinates.
            LastX = X: LastY = Y
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    RUNNING = False
    'Show the mouse if it was hidden at the program initiation.
    If RunMODE = "S" Then Call ShowMouse(True)

End Sub

'----------------------------------------------'
' TIMER is used to start the animation as this '
' will remove animation from the context of    '
' the form's RESIZE event, which would prevent '
' the form from unloading without an error.    '
'----------------------------------------------'
Private Sub tmrANI_Timer()

    RUNNING = True
    Call ANIMATE

End Sub

'-----------------------------------------------'
' The actual ANIMATION is performed using a     '
' timing loop instead of the Timer control as I '
' find that this is a more effective method of  '
' controlling the rate of display.              '
'-----------------------------------------------'
Private Sub ANIMATE()

    Dim BeginDrawTime As Double
    Dim MIN_DrawTime As Double
    
    'Once this routine has been invoked we no
    'longer need the Timer control running.
    tmrANI.Enabled = False
    
    'This calculation will yield a value of between 50 and 250 Milliseconds.
    MIN_DrawTime = 50 + 2 * (100 - CM.AniSpeed)
    
    Do While RUNNING
        BeginDrawTime = timeGetTime
        DoEvents: DRAW_1_FRAME: DoEvents
        Do While RUNNING And (timeGetTime - BeginDrawTime) < MIN_DrawTime
            DoEvents
        Loop
    Loop
    
    Unload Me

End Sub

Private Sub DRAW_1_FRAME()

    Static FacialMvmnt As Integer ' 0 = Mouth Opening
                                  ' 1 = Mouth Closing
                                  ' 2 = Eyes Glow
                                  ' 3 = Eyes dim a little
                                  ' 4 = Eyes Glow again
                                  ' 5 = Eyes completely dim
    Static MvmntPositn As Integer ' Indicates at what stage each of
                                  ' movement we are currently at.
    Static WordCnt As Integer     ' The number of words currently being
                                  ' shown on the form.
    Static CrntSentence As String ' The portion of the sentence which is
                                  ' currently being shown on the form.
    
    'This little dance does most of the movement coordination.
    Select Case FacialMvmnt
        Case 0: If MvmntPositn = 5 Then FacialMvmnt = 1: MvmntPositn = 4 Else MvmntPositn = MvmntPositn + 1
        Case 1: If MvmntPositn = 0 Then FacialMvmnt = 2 Else MvmntPositn = MvmntPositn - 2
        Case 2: If MvmntPositn = 5 Then FacialMvmnt = 3: MvmntPositn = 4 Else MvmntPositn = MvmntPositn + 1
        Case 3: If MvmntPositn = 3 Then FacialMvmnt = 4: MvmntPositn = 4 Else MvmntPositn = MvmntPositn - 1
        Case 4: If MvmntPositn = 5 Then FacialMvmnt = 5: MvmntPositn = 4 Else MvmntPositn = MvmntPositn + 1
        Case 5: If MvmntPositn = 0 Then FacialMvmnt = 0 Else MvmntPositn = MvmntPositn - 1
    End Select
    
    'If Charlie is "Talking"
    If FacialMvmnt < 2 Then
        'Stretch the appropraite image of Charlie's mouth in
        'an open position onto the image destination portion
        'of the form.
        StretchBlt Charles.hDC, PICT.LEFT, PICT.TOP, PICT.WIDTH, PICT.HEIGHT, _
                   picFACE.hDC, 0, 216 * MvmntPositn, 216, 216, vbSrcCopy
        'If the mouth is shut then determine if there are any
        'more words to be said, then if so we'll set the facial
        'movemnt to opening the mouth again.
        If FacialMvmnt = 1 And MvmntPositn = 0 And WordCnt < MaxWords Then FacialMvmnt = 0
    'If Charlie's eyes are glowing.
    Else
        'Stretch the appropraite image of Charlie's eyes
        'glowing onto the image destination portion of
        'the form.
        StretchBlt Charles.hDC, PICT.LEFT, PICT.TOP, PICT.WIDTH, PICT.HEIGHT, _
                   picFACE.hDC, 216, 216 * MvmntPositn, 216, 216, vbSrcCopy
        'At the end of the sequence we will clear the
        'text from the form by copying back the image
        'of the background we made at program initiation.
        If FacialMvmnt = 5 And MvmntPositn = 0 Then
            BitBlt Charles.hDC, TEXT.LEFT, TEXT.TOP, TEXT.WIDTH, TEXT.HEIGHT, picGRAD.hDC, 0, 0, vbSrcCopy
        End If
    End If
    
    'When Charlie's mouth is wide open, display the next word in the sentence.
    If FacialMvmnt = 0 And MvmntPositn = 5 Then
        
        'Add another word into the current sentence being displayed.
        If WordCnt = MaxWords Then WordCnt = 1 Else WordCnt = WordCnt + 1
        If WordCnt = 1 Then
            CrntSentence = TextWord(0)
        Else
            CrntSentence = CrntSentence + " " + TextWord(WordCnt - 1)
        End If
        
        'Load the saved image of the text destination area into picIMAG
        BitBlt picIMAG.hDC, 0, 0, TEXT.WIDTH, TEXT.HEIGHT, picGRAD.hDC, 0, 0, vbSrcCopy
        'Write the current sentence in picTEXT in Black on White
        picTEXT.BackColor = vbWhite: picTEXT.ForeColor = vbBlack
        RtnVal = DrawText(picTEXT.hDC, CrntSentence, -1, rctTEXT, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER)
        'Make a cutout for the text into the background in which it will sit.
        BitBlt picIMAG.hDC, 0, 0, TEXT.WIDTH, TEXT.HEIGHT, picTEXT.hDC, 0, 0, vbSrcAnd
        'Write the current sentence in picTEXT in White on Black
        picTEXT.BackColor = vbBlack: picTEXT.ForeColor = vbWhite
        RtnVal = DrawText(picTEXT.hDC, CrntSentence, -1, rctTEXT, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER)
        'Add the text gradient colours into the white text just created.
        BitBlt picTEXT.hDC, 0, 0, TEXT.WIDTH, TEXT.HEIGHT, picBACK.hDC, 0, 0, vbSrcAnd
        'Slot the gradiated text image into the background cutout.
        BitBlt picIMAG.hDC, 0, 0, TEXT.WIDTH, TEXT.HEIGHT, picTEXT.hDC, 0, 0, vbSrcPaint
        'Load the image of the gradiated text in its gradiated
        'background into the position on the form from where
        'the background image was taken.
        BitBlt Charles.hDC, TEXT.LEFT, TEXT.TOP, TEXT.WIDTH, TEXT.HEIGHT, picIMAG.hDC, 0, 0, vbSrcCopy
        
    End If

End Sub
