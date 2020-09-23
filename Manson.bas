Attribute VB_Name = "Manson"
Option Explicit

    Public RunMODE As String
    
    Type Fill_Options
        PictureFill As Boolean
        FillStyle As Integer
        ColourStyle As Integer
        GradColour(1 To 2) As OLE_COLOR
        ColourDrctn As Integer
        PictureType As Integer
        PictureFit As Integer
        PictureAddr As String
    End Type
        
    Type SS_Settings
        SavrText As String
        BelowPic As Boolean
        AniSpeed As Integer
        FontName As String
        FontBold As Boolean
        FontItlc As Boolean
        FontUndr As Boolean
        TEXT As Fill_Options
        BACK As Fill_Options
    End Type
    Public CM As SS_Settings
    
    Public TextWord() As String
    Public MaxWords   As Integer

    Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
    
    Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

    Public Const DT_CENTER = &H1
    Public Const DT_VCENTER = &H4
    Public Const DT_SINGLELINE = &H20
    
    Type RECT
        LEFT As Long
        TOP As Long
        Right As Long
        Bottom As Long
    End Type
        
    Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Type POINTAPI
        X As Long
        Y As Long
    End Type
    
    Public Const SWP_NOACTIVATE = &H10
    Public Const SWP_NOZORDER = &H4
    Public Const SWP_SHOWWINDOW = &H40
    
    Public Const HWND_TOP = 0
    Public Const HWND_TOPMOST = -1
    
    Public Const WS_CHILD = &H40000000
    Public Const GWL_HWNDPARENT = (-8)
    Public Const GWL_STYLE = (-16)
    
    Public Const STRETCH_ANDSCANS = 1
    Public Const STRETCH_ORSCANS = 2
    Public Const STRETCH_DELETESCANS = 3
    Public Const STRETCH_HALFTONE = 4
    
    Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Declare Function GetDesktopWindow Lib "user32" () As Long
    Declare Function PaintDesktop Lib "user32" (ByVal hDC As Long) As Long
    Declare Function GetStretchBltMode Lib "gdi32" (ByVal hDC As Long) As Long
    Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
    Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
    Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

    Public PNTR As Integer 'For general use as a counter
    Public RtnVal As Long  'For general use Returning Values from API calls
    
    Public DeskTop_hWnd As Long 'Used for the Desktop's hWnd
    Public DeskTop_DC As Long   'Used for the Desktop's Device Context


'----------------------------------------------------'
' This procedure is that which is initially invoked. '
'----------------------------------------------------'
Public Sub Main()

    Dim SSargs As String
    
    'Retrieve the startup command.
    SSargs = UCase(Command())
    If LEFT(SSargs, 1) = "/" Then
        RunMODE = Mid(SSargs, 2, 1)
    Else
        RunMODE = LEFT(SSargs, 1)
    End If
    
    'Act upon which command was received.
    Select Case RunMODE
        Case "":  Charles.Show vbModal
        Case "C": Settings.Show vbModal
        Case "S": Call ShowScreenSaver
        Case "P": Call PreviewWindow(SSargs)
        Case "A": MsgBox "No PASSWORD protection is available.", _
                          vbExclamation, _
                         "CHARLIE Password: [ WARNING ]"
        Case Else: End
    End Select

End Sub

'-----------------------------------------------------'
' Start an instance of the Form in Screen Saver mode. '
'-----------------------------------------------------'
Private Sub ShowScreenSaver()

    Const SSvrName As String = "S: Charlie II Screen Saver"
    
    'Stop if there is already an instance of the screen saver mode.
    If App.PrevInstance Then
        If FindWindow(vbNullString, SSvrName) Then End
    End If
    
    'Set this instances caption so other instances can find this one.
    Charles.Caption = SSvrName
    
    'Load the form and display it as the topmost form.
    Load Charles
    SetWindowPos Charles.hWnd, HWND_TOPMOST, _
                 0&, 0&, Screen.WIDTH, Screen.HEIGHT, _
                 SWP_SHOWWINDOW
    
End Sub

'------------------------------------------------'
' Start an instance of the Form in Preveiw mode. '
'------------------------------------------------'
Private Sub PreviewWindow(SScmmnd As String)

    Const PrvwName As String = "P: Charlie II Screen Saver"
    Dim PPhWnd As Double
    Dim PPrect As RECT
    Dim FWstyle As Long

    'Extract the hWnd of the Preview Pane passed from
    'Display Properties when it started this programme.
    For PNTR = 1 To Len(SScmmnd)
        If Val(Mid(SScmmnd, PNTR)) > 0 Then
            PPhWnd = CDbl(Mid(SScmmnd, PNTR))
            Exit For
        End If
    Next PNTR
    
    'Get the dimensions of the Preview Pane.
    GetClientRect PPhWnd, PPrect
    
    'Load the Form and set it's caption.
    Load Charles
    Charles.Caption = PrvwName
    
    'Get the Form's window style and alter it to be a child window,
    'then set the Preview Pane to be the Form's parent.
    FWstyle = GetWindowLong(Charles.hWnd, GWL_STYLE)
    SetWindowLong Charles.hWnd, GWL_STYLE, FWstyle Or WS_CHILD
    SetParent Charles.hWnd, PPhWnd
    
    'Show the Screen Saver Form within the Preview Pane area.
    SetWindowPos Charles.hWnd, HWND_TOP, _
                 0&, 0&, PPrect.Right, PPrect.Bottom, _
                 SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW

End Sub

'--------------------------'
' Show or Hide the cursor. '
'--------------------------'
Public Sub ShowMouse(Optional Visible As Boolean = True)

    Dim Rtn As Integer

    If Visible Then
        Do: Rtn = ShowCursor(True): Loop Until Rtn >= 0
    Else
        Do: Rtn = ShowCursor(False): Loop Until Rtn < 0
    End If

End Sub

'---------------------------------|
' Get settings from the Registry. |
'---------------------------------|
Public Sub GetValues()

    CM.SavrText = GetSetting("Charlie II", "General", "Text", "Charlie II Screen Saver")
    CM.BelowPic = GetSetting("Charlie II", "General", "Text Below Picture", True)
    CM.AniSpeed = GetSetting("Charlie II", "General", "Animation Speed", 50)
    
    CM.FontName = GetSetting("Charlie II", "Font", "Name", "Times New Roman")
    CM.FontBold = GetSetting("Charlie II", "Font", "Bold", True)
    CM.FontItlc = GetSetting("Charlie II", "Font", "Italic", False)
    CM.FontUndr = GetSetting("Charlie II", "Font", "Underline", False)
    
    CM.TEXT.PictureFill = GetSetting("Charlie II", "Text Fill", "Picture Fill", False)
    CM.TEXT.FillStyle = GetSetting("Charlie II", "Text Fill", "Fill Style", 0)
    CM.TEXT.ColourStyle = GetSetting("Charlie II", "Text Fill", "Gradient Style", 1)
    CM.TEXT.GradColour(1) = GetSetting("Charlie II", "Text Fill", "Gradient Colour (1)", vbYellow)
    CM.TEXT.GradColour(2) = GetSetting("Charlie II", "Text Fill", "Gradient Colour (2)", vbRed)
    CM.TEXT.ColourDrctn = GetSetting("Charlie II", "Text Fill", "Gradient Direction", 1)
    CM.TEXT.PictureType = GetSetting("Charlie II", "Text Fill", "Picture Type", 0)
    CM.TEXT.PictureFit = GetSetting("Charlie II", "Text Fill", "Picture Fit", 0)
    CM.TEXT.PictureAddr = GetSetting("Charlie II", "Text Fill", "Picture Address", "C:\Windows\Red Blocks.bmp")
    
    CM.BACK.PictureFill = GetSetting("Charlie II", "Back Fill", "Picture Fill", False)
    CM.BACK.FillStyle = GetSetting("Charlie II", "Back Fill", "Fill Style", 0)
    CM.BACK.ColourStyle = GetSetting("Charlie II", "Back Fill", "Gradient Style", 3)
    CM.BACK.GradColour(1) = GetSetting("Charlie II", "Back Fill", "Gradient Colour (1)", vbCyan)
    CM.BACK.GradColour(2) = GetSetting("Charlie II", "Back Fill", "Gradient Colour (2)", vbBlue)
    CM.BACK.ColourDrctn = GetSetting("Charlie II", "Back Fill", "Gradient Direction", 1)
    CM.BACK.PictureType = GetSetting("Charlie II", "Back Fill", "Picture Type", 0)
    CM.BACK.PictureFit = GetSetting("Charlie II", "Back Fill", "Picture Fit", 0)
    CM.BACK.PictureAddr = GetSetting("Charlie II", "Back Fill", "Picture Address", "C:\Windows\Clouds.bmp")
    
    'Break the text sentence down into individual words.
    Dim Strt As Long, Pstn As Long
    Dim WCnt As Integer
    Strt = 1: Pstn = 0: WCnt = 0
    Do
        Pstn = InStr(Strt, CM.SavrText, " ")
        If Pstn = 0 Then Pstn = Len(CM.SavrText) + 1
        ReDim Preserve TextWord(WCnt)
        TextWord(WCnt) = Mid(CM.SavrText, Strt, Pstn - Strt)
        Strt = Pstn + 1
        WCnt = WCnt + 1
    Loop Until Strt >= Len(CM.SavrText)
    MaxWords = WCnt

End Sub

'-----------------------------------|
' Write settings into the Registry. |
'-----------------------------------|
Public Sub PutValues()

    SaveSetting "Charlie II", "General", "Text", CM.SavrText
    SaveSetting "Charlie II", "General", "Text Below Picture", CM.BelowPic
    SaveSetting "Charlie II", "General", "Animation Speed", CM.AniSpeed
    
    SaveSetting "Charlie II", "Font", "Name", CM.FontName
    SaveSetting "Charlie II", "Font", "Bold", CM.FontBold
    SaveSetting "Charlie II", "Font", "Italic", CM.FontItlc
    SaveSetting "Charlie II", "Font", "Underline", CM.FontUndr
    
    SaveSetting "Charlie II", "Text Fill", "Picture Fill", CM.TEXT.PictureFill
    SaveSetting "Charlie II", "Text Fill", "Fill Style", CM.TEXT.FillStyle
    SaveSetting "Charlie II", "Text Fill", "Gradient Style", CM.TEXT.ColourStyle
    SaveSetting "Charlie II", "Text Fill", "Gradient Colour (1)", CM.TEXT.GradColour(1)
    SaveSetting "Charlie II", "Text Fill", "Gradient Colour (2)", CM.TEXT.GradColour(2)
    SaveSetting "Charlie II", "Text Fill", "Gradient Direction", CM.TEXT.ColourDrctn
    SaveSetting "Charlie II", "Text Fill", "Picture Type", CM.TEXT.PictureType
    SaveSetting "Charlie II", "Text Fill", "Picture Fit", CM.TEXT.PictureFit
    SaveSetting "Charlie II", "Text Fill", "Picture Address", CM.TEXT.PictureAddr
    
    SaveSetting "Charlie II", "Back Fill", "Picture Fill", CM.BACK.PictureFill
    SaveSetting "Charlie II", "Back Fill", "Fill Style", CM.BACK.FillStyle
    SaveSetting "Charlie II", "Back Fill", "Gradient Style", CM.BACK.ColourStyle
    SaveSetting "Charlie II", "Back Fill", "Gradient Colour (1)", CM.BACK.GradColour(1)
    SaveSetting "Charlie II", "Back Fill", "Gradient Colour (2)", CM.BACK.GradColour(2)
    SaveSetting "Charlie II", "Back Fill", "Gradient Direction", CM.BACK.ColourDrctn
    SaveSetting "Charlie II", "Back Fill", "Picture Type", CM.BACK.PictureType
    SaveSetting "Charlie II", "Back Fill", "Picture Fit", CM.BACK.PictureFit
    SaveSetting "Charlie II", "Back Fill", "Picture Address", CM.BACK.PictureAddr

End Sub

'--------------------------------------------------------'
' This procedure will draw a colour graduated background '
' onto either a Form or a Picture Box Control.           '
'--------------------------------------------------------'
Sub GradFill(FillObj As Object, _
             Optional ByVal StartRGB As OLE_COLOR = vbRed, _
             Optional ByVal FinalRGB As OLE_COLOR = vbBlack, _
             Optional ByRef FillStyle As Integer = 1, _
             Optional ByRef Orientation As Integer = 1)

    Dim StartRED As Integer, StartGRN As Integer, StartBLU As Integer
    Dim FinalRED As Integer, FinalGRN As Integer, FinalBLU As Integer
    Dim CrrntRED As Integer, CrrntGRN As Integer, CrrntBLU As Integer, CrrntRGB As OLE_COLOR
    Dim SavedScaleMode As Integer, SavedPenSize As Integer
    Dim Reps As Integer, Cntr As Integer, CTFctr As Double
    
    Select Case FillStyle
        Case 0
            'No Gradient Fill required, fill plain background then exit.
            FillObj.BackColor = StartRGB
            FillObj.Refresh
            Exit Sub
        Case 1 To 5
            'No action required here.
        Case Else
            'Display a Message Box about there being an invalid value for style.
            MsgBox "P R O G R A M   E R R O R   -   ( non fatal )" & vbCrLf & vbCrLf & _
                   "Invalid value provided for parameter ""FillStyle""." & vbCrLf & vbCrLf & _
                   "The valid values are integers in the range: 0 to 5." & vbCrLf & vbCrLf & _
                   "No Gradient Fill will be drawn on " & FillObj.Name & ".", _
                   vbExclamation + vbOKOnly, _
                   "Sub: GradientFill"
            Exit Sub
    End Select
     
    Select Case Orientation
        Case 1, 3, 5
            'No action required, all is OK
        Case 2, 4, 6
            'If the direction is an even number then
            'Start and End colours are reversed.
            CrrntRGB = StartRGB
            StartRGB = FinalRGB
            FinalRGB = CrrntRGB
        Case Else
            'Display a Message Box about there being an invalid value for orientation.
            MsgBox "P R O G R A M   E R R O R   -   ( non fatal )" & vbCrLf & vbCrLf & _
                   "Invalid value provided for parameter ""Orientation""." & vbCrLf & vbCrLf & _
                   "The valid values are integers in the range: 1 to 6." & vbCrLf & vbCrLf & _
                   "No Gradient Fill will be drawn on " & FillObj.Name & ".", _
                   vbExclamation + vbOKOnly, _
                   "Sub: GradientFill"
            Exit Sub
    End Select
    
    'Seperate out the R,G,B components of the Starting colour.
    StartRED = StartRGB And &HFF
    StartGRN = Int((StartRGB - StartRED) / 256) And &HFF
    StartBLU = Int((((StartRGB - StartRED) / 256) - StartGRN) / 256) And &HFF
    
    'Seperate out the R,G,B components of the Ending colour.
    FinalRED = FinalRGB And &HFF
    FinalGRN = Int((FinalRGB - FinalRED) / 256) And &HFF
    FinalBLU = Int((((FinalRGB - FinalRED) / 256) - FinalGRN) / 256) And &HFF
    
    'If the object is not in scale pixels then set it to scale pixels
    SavedScaleMode = FillObj.ScaleMode
    If FillObj.ScaleMode <> 3 Then FillObj.ScaleMode = 3
    'If the object's pen size is not set to 1 then set it to 1.
    SavedPenSize = FillObj.DrawWidth
    If FillObj.DrawWidth <> 1 Then FillObj.DrawWidth = 1
    
    'Determine the number of repetitions of a line drawing
    'that will be required to fill the drawing object.
    Select Case FillStyle
        Case 1:      Reps = FillObj.ScaleHeight
        Case 2:      Reps = FillObj.ScaleWidth
        Case 3 To 4: Reps = 2 * (FillObj.ScaleHeight + FillObj.ScaleWidth)
        Case 5:      Reps = IIf(FillObj.ScaleWidth > FillObj.ScaleHeight, _
                                FillObj.ScaleWidth / 2, _
                                FillObj.ScaleHeight / 2)
    End Select
    
    For Cntr = 0 To Reps
    
        'Compute the amount of Colour Transition from the
        'first to second colour for each line being drawn.
        Select Case Orientation
            'For Single Colour Transitions
            Case 1 To 2
                CTFctr = Cntr / Reps
            'For Double Colour Transitions
            Case 3 To 4
                Select Case Cntr
                    Case Is <= Reps * 0.5:  CTFctr = Cntr / Reps * 2
                    Case Is <= Reps:        CTFctr = (Reps - Cntr) / Reps * 2
                End Select
            'For Quadruple Colour Transitions
            Case 5 To 6
                Select Case Cntr
                    Case Is <= Reps * 0.25: CTFctr = Cntr / Reps * 4
                    Case Is <= Reps * 0.5:  CTFctr = (Reps / 2 - Cntr) / Reps * 4
                    Case Is <= Reps * 0.75: CTFctr = (Cntr - Reps / 2) / Reps * 4
                    Case Is <= Reps:        CTFctr = (Reps - Cntr) / Reps * 4
                End Select
        End Select
        
        'Ensure that Colour Transition factor has a value in the range of 0 to 1.
        If CTFctr > 1 Then CTFctr = 1
        If CTFctr < 0 Then CTFctr = 0
        
        'Calculate the RGB of the colour to be drawn.
        CrrntRED = StartRED + CInt(CTFctr * (FinalRED - StartRED))
        CrrntGRN = StartGRN + CInt(CTFctr * (FinalGRN - StartGRN))
        CrrntBLU = StartBLU + CInt(CTFctr * (FinalBLU - StartBLU))
        CrrntRGB = RGB(CrrntRED, CrrntGRN, CrrntBLU)
        
        'Draw a line of the appropriate style in the calculated colour.
        Select Case FillStyle
            'Style = Horizontal
            Case 1: FillObj.Line (0, Cntr)-(FillObj.ScaleWidth, Cntr), CrrntRGB
            'Style = Vertical
            Case 2: FillObj.Line (Cntr, 0)-(Cntr, FillObj.ScaleHeight), CrrntRGB
            'Style = Diagonal Up
            Case 3: FillObj.Line (0, FillObj.ScaleHeight / Reps * 2 * Cntr)- _
                                 (FillObj.ScaleWidth / Reps * 2 * Cntr, -1), _
                                  CrrntRGB
            'Style = Diagonal Down
            Case 4: FillObj.Line (FillObj.ScaleWidth, FillObj.ScaleHeight / Reps * 2 * Cntr)- _
                                 (FillObj.ScaleWidth - FillObj.ScaleWidth / Reps * 2 * Cntr, -1), _
                                  CrrntRGB
            'Style = Square
            Case 5: FillObj.Line (FillObj.ScaleWidth / 2 / Reps * Cntr, _
                                  FillObj.ScaleHeight / 2 / Reps * Cntr)- _
                                 (FillObj.ScaleWidth - FillObj.ScaleWidth / 2 / Reps * Cntr, _
                                  FillObj.ScaleHeight - FillObj.ScaleHeight / 2 / Reps * Cntr), _
                                  CrrntRGB, B
                    If Cntr = Reps Then FillObj.PSet (FillObj.ScaleWidth / 2, _
                                                      FillObj.ScaleHeight / 2), CrrntRGB
        End Select
        
    Next Cntr
    
    'If the object's scale wasn't originally in pixels
    'then set it back to its original scale.
    If FillObj.ScaleMode <> SavedScaleMode Then FillObj.ScaleMode = SavedScaleMode
    'If the object's pen size has been changed then
    'set it back to its original size.
    If FillObj.DrawWidth <> SavedPenSize Then FillObj.DrawWidth = SavedPenSize

End Sub

'----------------------------------------------------------------'
' This procedure will draw a picture from a File into a Form     '
' or a Picture Box Control using a selected type of picture fit. '
'----------------------------------------------------------------'
Sub PictFill(DestObj As Object, _
             WorkObj As Object, _
             ByVal PictFile As String, _
             Optional ByRef FitType As Integer = 0)

    'Clear the Destination Object
    DestObj.Cls
    
    'Load the User Selected Image into the Work Object
    Set WorkObj.Picture = LoadPicture(PictFile)
    
    'Save the Stretch Mode of the Destination Object and
    'then alter it's Stretch Mode to "HalfTone".
    Dim SaveScan_DestObj As Long
    SaveScan_DestObj = GetStretchBltMode(DestObj.hDC)
    RtnVal = SetStretchBltMode(DestObj.hDC, STRETCH_HALFTONE)
    
    'Load the picture into the Destination Object using the given "Fit".
    Select Case FitType
        
        'CENTRE the picture in the Destination Object
        Case 0
            RtnVal = BitBlt(DestObj.hDC, _
                            (DestObj.WIDTH - WorkObj.WIDTH) / 2, _
                            (DestObj.HEIGHT - WorkObj.HEIGHT) / 2, _
                            WorkObj.WIDTH, WorkObj.HEIGHT, _
                            WorkObj.hDC, 0, 0, vbSrcCopy)
        
        'TILE the picture into the Destination Object
        Case 1
            Dim CntX As Integer, CntY As Integer
            For CntX = 0 To DestObj.WIDTH Step WorkObj.WIDTH
            For CntY = 0 To DestObj.HEIGHT Step WorkObj.HEIGHT
                RtnVal = BitBlt(DestObj.hDC, _
                                CntX, CntY, WorkObj.WIDTH, WorkObj.HEIGHT, _
                                WorkObj.hDC, 0, 0, vbSrcCopy)
            Next CntY: Next CntX
        
        'STRETCH the picture so it fits into the Destination Object
        Case 2
            RtnVal = StretchBlt(DestObj.hDC, 0, 0, 160, 120, _
                                WorkObj.hDC, 0, 0, _
                                WorkObj.WIDTH, _
                                WorkObj.HEIGHT, vbSrcCopy)
        
        'ZOOM the picture In or Out so that it fits into the
        'Destination Object while maintaining it's Aspect Ratio.
        Case 3
            Dim Fctr As Double
            Fctr = IIf((WorkObj.WIDTH / DestObj.WIDTH) < _
                       (WorkObj.HEIGHT / DestObj.HEIGHT), _
                       (DestObj.WIDTH / WorkObj.WIDTH), _
                       (DestObj.HEIGHT / WorkObj.HEIGHT))
            RtnVal = StretchBlt(DestObj.hDC, _
                                (DestObj.WIDTH - Fctr * WorkObj.WIDTH) / 2, _
                                (DestObj.HEIGHT - Fctr * WorkObj.HEIGHT) / 2, _
                                Fctr * WorkObj.WIDTH, _
                                Fctr * WorkObj.HEIGHT, _
                                WorkObj.hDC, 0, 0, WorkObj.WIDTH, WorkObj.HEIGHT, _
                                vbSrcCopy)
        
        'Display a Message Box about there being an invalid
        'value for the type of Picture Fit.
        Case Else
            MsgBox "P R O G R A M   E R R O R   -   ( non fatal )" & vbCrLf & vbCrLf & _
                   "Invalid value provided for parameter ""FitType""." & vbCrLf & vbCrLf & _
                   "The valid values are integers in the range: 1 to 4." & vbCrLf & vbCrLf & _
                   "No Picture will be drawn into " & DestObj.Name & ".", _
                   vbExclamation + vbOKOnly, _
                   "Sub: PictFill"
        
    End Select
    
    'Restore the Stretch Mode of the Destination Object.
    RtnVal = SetStretchBltMode(DestObj.hDC, SaveScan_DestObj)

End Sub

