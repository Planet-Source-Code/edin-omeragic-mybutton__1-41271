VERSION 5.00
Begin VB.UserControl MyButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   DefaultCancel   =   -1  'True
   FillStyle       =   0  'Solid
   PropertyPages   =   "MyButton.ctx":0000
   ScaleHeight     =   75
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   271
   Begin VB.Timer MyTimer 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   360
      Top             =   360
   End
End
Attribute VB_Name = "MyButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
                         'FREE SOURCE CODE'
'==================================================================
'Developed by edin omeragic
'From: Bosnia and Hercegovina
'e-mail: edoo_ba@hotmail.com
'Date: (20.11 - 3.12) 2002 godine
'==================================================================
'                 Joj ja sam pijan ko ce me popeti                '
'==================================================================
'if you dont like this you may copy it on flopy and throw it away ;},
    'If you Like it Then please vote on planetsourcecode
'and also search for:
' - iMenu (old project but good) or
' - iList (cool list)
'==================================================================
'DrawButton(State)  - draws button (main function)
'DrawText(...)      - draws text (called from drawbutton)
'DrawPicture(...)   - draws picture
'DrawPictureDisabled - draws picture grayed
'TilePicture()       - tiles picture
'SetRect (left, top, right, bottom) as RECT 'makes rectangle on fly
'ModyfyRect(RECT,left,top,right,bottom) as RECT
'i.e.
'R = SetRect     (0,0,1,1)
'R = ModifyRect(R,1,1,1,1)
'R is            (1,1,2,2)
'==================================================================
'-for default skin, name the picture box "MyButtonDefSkin"
'-for changing skin in design time set property
' "SkinPictureName" same as picture box name
'==================================================================

'UPDATE 1.  (5-6.12.2002 - 10.51) (d.m.y format)
'more samples added
'some bug fixed
'if skin is not set DrawFrameControl api is used to draw button
'implemented gradient fill

'UPDATE 2.
'ToolTipText now works

'Update 3.
'Default i cancel
'Buttons group



Option Explicit

Const CONTROL_NAME = "MyButton"

'Default Property Values:
Const m_def_OwnerDraw = False
Const m_def_Group = 0
Const m_def_Pressed = 0
'Const m_def_Group = 0
Const m_def_DisableGradient = 0
Const m_def_DisableHover = False
Const m_def_TextAlign = vbCenter
Const m_def_PictureTColor = &HFF00FF
Const m_def_PicturePos = 0
Const m_def_TextColorDisabled2 = 0
Const m_def_DrawFocus = 0
Const m_def_DisplaceText = 0
'Const m_def_DownTextDX = 0
'Const m_def_DownTextDY = 0
'Const m_def_DisableHover = False
Const m_def_TextColorEnabled = 0
Const m_def_TextColorDisabled = 0
Const m_def_FillWithColor = True
Const m_def_SizeCW = 3
Const m_def_SizeCH = 3
Const m_def_Text = ""
'Property Variables:
Dim m_OwnerDraw As Boolean
Dim m_Group As Long
Dim m_Pressed As Boolean
'Dim m_Group As Variant
Dim m_DisableGradient As Boolean
Dim m_DisableHover As Boolean
Dim m_TextAlign As AlignmentConstants
Dim m_PictureTColor As Ole_Color
Dim m_PicturePos As Integer
Dim m_Picture As StdPicture
Dim m_TextColorDisabled2 As Ole_Color
Dim m_DrawFocus As Integer
Dim m_DisplaceText As Integer
'Dim m_DisableHover As Boolean
Dim m_TextColorEnabled As Ole_Color
Dim m_TextColorDisabled As Ole_Color
Dim m_FillWithColor As Boolean
Dim m_SizeCW As Long
Dim m_SizeCH As Long
Dim m_SkinPicture As PictureBox
Dim m_Text As String
Dim m_State As Integer
Dim m_HasFocus As Boolean
Dim m_BtnDown As Boolean
Dim m_SpcDown As Boolean
Dim m_SkinPictureName As String
Dim m_RightBtn As Boolean


Public Enum EnumPicturePos
    ppLeft
    ppTop
    ppBottom
    ppRight
    ppCenter
End Enum
Private Const DI_NORMAL As Long = &H3

Const BTN_NORMAL = 1
Const BTN_FOCUS = 2
Const BTN_HOVER = 3
Const BTN_DOWN = 4
Const BTN_DISABLED = 5
'Event Declarations:
Event OnDrawButton(ByVal State As Integer)
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MouseHover()
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event KeyPress(KeyAscii As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Public Enum EnumDrawTextFormat
    DT_BOTTOM = &H8
    DT_CALCRECT = &H400
    DT_CENTER = &H1
    DT_CHARSTREAM = 4
    DT_DISPFILE = 6
    DT_EXPANDTABS = &H40
    DT_EXTERNALLEADING = &H200
    DT_INTERNAL = &H1000
    DT_LEFT = &H0
    DT_METAFILE = 5
    DT_NOCLIP = &H100
    DT_NOPREFIX = &H800
    DT_PLOTTER = 0
    DT_RASCAMERA = 3
    DT_RASDISPLAY = 1
    DT_RASPRINTER = 2
    DT_RIGHT = &H2
    DT_SINGLELINE = &H20
    DT_TABSTOP = &H80
    DT_TOP = &H0
    DT_VCENTER = &H4
    DT_WORDBREAK = &H10
    DT_WORD_ELLIPSIS = &H40000
    DT_END_ELLIPSIS = 32768
    DT_PATH_ELLIPSIS = &H4000
    DT_EDITCONTROL = &H2000
    '===================
    DT_INCENTER = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
End Enum

Private Const SRCCOPY = &HCC0020
Private Const RGN_AND = 1

Private Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectClipPath Lib "gdi32" (ByVal hDC As Long, ByVal iMode As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function apiDrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function apiTranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As Ole_Color, ByVal palet As Long, Col As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
'Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function ReleaseCapture Lib "user32" () As Long
'replaced with timer control

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
'MY NOTE: TransparentBlt on Win98 leavs some garbage in memory...
'
'Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long

'for picture
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
'Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal DW As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
'never enough
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long


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
Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(1) As RGBQUAD
End Type

'for checking if api is supported
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'draws standard button
Private Const DFC_BUTTON = 4
Private Const DFCS_BUTTONPUSH = &H10
Private Const DFCS_PUSHED = &H200
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long


'for gradient fill
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

'for mouse (in/out) - ToolTipText doesnt work if SetCapture is used
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'personal constants
Const MISION_IMPOSSIBLE = -1 'why use high math to callculate area of triangle


'#############################################
'//GDI + SOMETHING ELSE#######################
Sub GradientFill(hDC As Long, Colors() As Long, x As Long, y As Long, W As Long, H As Long, bVertical As Boolean)
'    If (UBound(Colors) - LBound(Colors)) = 0 Then
'        Debug.Print "GradientFill::At least 2 colors required."
'        Exit Sub
'    End If
    
    Dim I As Long
    Dim J As Long
    
    Dim Col1 As Long
    Dim Col2 As Long
    
    Dim StepR As Single
    Dim StepG As Single
    Dim StepB As Single
    
    Dim R1 As Integer
    Dim G1 As Integer
    Dim B1 As Integer
    
    Dim R2 As Integer
    Dim G2 As Integer
    Dim B2 As Integer
    
    
    Dim Col As Long
    
    Dim pFrom As Long
    Dim pTo As Long
    Dim Cnt As Long
    
    Dim hPen As Long
    
    For I = 0 To UBound(Colors) - 1
        Col1 = Colors(I)
        Col2 = Colors(I + 1)
                
        If bVertical Then
            pFrom = I * H / (UBound(Colors))
            pTo = (I + 1) * H / (UBound(Colors))
        Else
            pFrom = I * W / (UBound(Colors))
            pTo = (I + 1) * W / (UBound(Colors))
        End If
          
        
        GetRgb Colors(I), R1, G1, B1
        GetRgb Colors(I + 1), R2, G2, B2
        
        StepR = (R2 - R1) / (pTo - pFrom)
        StepG = (G2 - G1) / (pTo - pFrom)
        StepB = (B2 - B1) / (pTo - pFrom)
        
        For Cnt = pFrom To pTo
            R2 = R1 + (Cnt - pFrom) * StepR
            G2 = G1 + (Cnt - pFrom) * StepG
            B2 = B1 + (Cnt - pFrom) * StepB
            'make new pen
            hPen = CreatePen(0, 1, RGB(R2, G2, B2))
            'select new pen and delete old
            Call DeleteObject(SelectObject(hDC, hPen))
            'Draw vertical line
            If bVertical Then
                MoveToEx hDC, x, y + Cnt, ByVal 0
                LineTo hDC, x + W, y + Cnt
            Else
                MoveToEx hDC, Cnt + x, y, ByVal 0
                LineTo hDC, Cnt + x, y + H
            End If
        Next
    Next
End Sub



Private Sub TransBlt(ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, _
            ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, _
            ByVal xSrc As Long, ByVal ySrc As Long, ByVal clrMask As Ole_Color)
    
'one check to see if GdiTransparentBlt is supported
'better way to check if function is suported is using LoadLibrary and GetProcAdress
'than using GetVersion or GetVersionEx
'=====================================================
    Dim Lib As Long
    Dim ProcAdress As Long
    Dim lMaskColor As Long
    lMaskColor = TranslateColor(clrMask)
    Lib = LoadLibrary("gdi32.dll")
    '--------------------->make sure to specify corect name for function
    ProcAdress = GetProcAddress(Lib, "GdiTransparentBlt")
    FreeLibrary Lib
    If ProcAdress <> 0 Then
        'works on XP
        GdiTransparentBlt hdcDest, xDest, yDest, nWidth, nHeight, hdcSrc, xSrc, ySrc, nWidth, nHeight, lMaskColor
        'Debug.Print "Gdi transparent blt"
        Exit Sub 'make it short
    End If
'=====================================================
    Const DSna              As Long = &H220326
    Dim hdcMask             As Long
    Dim hdcColor            As Long
    Dim hbmMask             As Long
    Dim hbmColor            As Long
    Dim hbmColorOld         As Long
    Dim hbmMaskOld          As Long
    Dim hdcScreen           As Long
    Dim hdcScnBuffer        As Long
    Dim hbmScnBuffer        As Long
    Dim hbmScnBufferOld     As Long
    

   hdcScreen = UserControl.hDC
   
   lMaskColor = TranslateColor(clrMask)
   hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
   hdcScnBuffer = CreateCompatibleDC(hdcScreen)
   hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)

   BitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcDest, xDest, yDest, vbSrcCopy

   hbmColor = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
   hbmMask = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)

   hdcColor = CreateCompatibleDC(hdcScreen)
   hbmColorOld = SelectObject(hdcColor, hbmColor)
    
   Call SetBkColor(hdcColor, GetBkColor(hdcSrc))
   Call SetTextColor(hdcColor, GetTextColor(hdcSrc))
   Call BitBlt(hdcColor, 0, 0, nWidth, nHeight, hdcSrc, xSrc, ySrc, vbSrcCopy)

   hdcMask = CreateCompatibleDC(hdcScreen)
   hbmMaskOld = SelectObject(hdcMask, hbmMask)

   SetBkColor hdcColor, lMaskColor
   SetTextColor hdcColor, vbWhite
   BitBlt hdcMask, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcCopy
 
   SetTextColor hdcColor, vbBlack
   SetBkColor hdcColor, vbWhite
   BitBlt hdcColor, 0, 0, nWidth, nHeight, hdcMask, 0, 0, DSna
   BitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcMask, 0, 0, vbSrcAnd
   BitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcPaint
   BitBlt hdcDest, xDest, yDest, nWidth, nHeight, hdcScnBuffer, 0, 0, vbSrcCopy
     
     'clear
   DeleteObject SelectObject(hdcColor, hbmColorOld)
   DeleteDC hdcColor
   DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
   DeleteDC hdcScnBuffer
   DeleteObject SelectObject(hdcMask, hbmMaskOld)
   
   DeleteDC hdcMask
   'ReleaseDC 0, hdcScreen
End Sub

Private Function GetRgbQuad(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte) As RGBQUAD
    With GetRgbQuad
        .rgbBlue = B
        .rgbGreen = G
        .rgbRed = R
    End With
End Function
Private Function DrawPictureDisabled(ByVal p As StdPicture, x As Long, y As Long, _
                 W As Long, H As Long, _
                 Optional ColHighlight As Long = vb3DHighlight, _
                 Optional ColShadow As Long = vb3DShadow)
                 
    'Dim MemDC As Long
    'Dim MyBmp As Long
    Dim cShadow As Long
    Dim cHiglight As Long
    Dim ColPal(0 To 1) As RGBQUAD
    Dim rgbBlack As RGBQUAD
    Dim rgbWhite As RGBQUAD
    Dim BI As BITMAPINFO
    Dim hDC As Long
    Dim hPicDc As Long
    Dim hPicBmp As Long
    hDC = UserControl.hDC
    
    cHiglight = TranslateColor(ColHighlight)
    cShadow = TranslateColor(ColShadow) 'bug fixed (ColShadow - vb3DShadow)
    
    'rgbBlack = GetRgbQuad(0, 0, 0)
    rgbWhite = GetRgbQuad(255, 255, 255)
    
    With BI.bmiHeader
        .biSize = 40 'size of bmiHeader structure
        .biHeight = -H
        .biWidth = W
        .biPlanes = 1
        .biCompression = 0 'BI_RGB
        .biClrImportant = 0
        .biBitCount = 1 'monohrome bitmap
    End With
    
    'color palete
    With BI
        .bmiColors(0) = rgbBlack
        .bmiColors(1) = rgbWhite
    End With
    
    Dim hMonoSec As Long
    Dim pBits As Long
    Dim hdcMono As Long
    
    hMonoSec = CreateDIBSection(hDC, BI, 0, pBits, 0&, 0&)
    'Debug.Print "MonoSec:"; hMonoSec
    hdcMono = CreateCompatibleDC(hDC)
    SelectObject hdcMono, hMonoSec
    Call DeleteObject(hMonoSec)
    
    'create dc for picture
    hPicDc = CreateCompatibleDC(hDC)
    If p.Type = vbPicTypeIcon Then
        hPicBmp = CreateCompatibleBitmap(hDC, W, H)
        SelectObject hPicDc, hPicBmp
        DeleteObject hPicBmp
        ClearRect hPicDc, SetRect(0, 0, W, H), vbWhite
        DrawIconEx hPicDc, 0, 0, p.handle, W, H, 0, 0, DI_NORMAL
        'Debug.Print "DRAW ICON"
    ElseIf p.Type = vbPicTypeBitmap Then
        SelectObject hPicDc, p.handle
    End If
    '--^ supports only GIF, JPEG, BMP and ICO files
        
    
    'copy  hPicDc to hdcMono, (convert color to black and white)
    BitBlt hdcMono, 0, 0, W, H, hPicDc, 0, 0, SRCCOPY
    
    DeleteDC (hPicDc)
    
    Dim R As Integer, G As Integer, B   As Integer
    GetRgb cHiglight, R, G, B
    
    'change black color in palete to highlight(r,g,b) color
    ColPal(0) = GetRgbQuad(R, G, B)
    ColPal(1) = rgbBlack    'change white color in palete to black color
    
    SetDIBColorTable hdcMono, 0, 2, ColPal(0)   'set new palete
    RealizePalette hdcMono                      'update it
    'BitBlt Me.hdc, 1, 1, W, H, hdcMono,  0, 0, SRCCOPY
      
    'transparent blit to dest hDC using black as transparent colour
    'x+1 and y+1 - moves down and left for 1 pixel
    TransBlt hDC, x + 1, y + 1, W, H, hdcMono, 0, 0, 0
    'the same as drawing disabled text
    
    'get rgb components of shadow color
    GetRgb cShadow, R, G, B
    'change black color to shadow color in palete
    ColPal(0) = GetRgbQuad(R, G, B)
    ColPal(1) = rgbWhite 'change back to white
    
    'set new palete
    SetDIBColorTable hdcMono, 0, 2, ColPal(0)
    RealizePalette hdcMono ' then update
    
    'transparent blit do dest hdc using white color as transparent
    TransBlt hDC, x, y, W, H, hdcMono, 0, 0, RGB(255, 255, 255)
    
    'BitBlt Me.hDC, 0, 0, W, H, hdcMono, 0, 0, SRCCOPY
    
    
    DeleteObject (hdcMono)  'bug fixed (it was commented)
   
End Function
Sub GetRgb(Color As Long, R As Integer, G As Integer, B As Integer)
       R = Color And 255            'clear bites from 9 to 32
       G = (Color \ 256) And 255    'shift right 8 bits and clear
       B = (Color \ 65536) And 255  'shift 16 bits and clear for any case
End Sub

Private Function GetBmpSize(Bmp As StdPicture, W As Long, H As Long) As Long
    W = ScaleX(Bmp.Width, vbHimetric, vbPixels)
    H = ScaleY(Bmp.Height, vbHimetric, vbPixels)
End Function

Private Sub DrawPicture(hDC As Long, p As StdPicture, x As Long, y As Long, W As Long, H As Long, TOleCol As Long)
    
    'check picture format
    If p.Type = vbPicTypeIcon Then
        DrawIconEx hDC, x, y, p.handle, W, H, 0, 0, DI_NORMAL
        Exit Sub
    End If
    
    'creting dc with the same format as screen dc
    Dim MemDC As Long
    MemDC = CreateCompatibleDC(hDC)
    
    'select a picture into memdc
    SelectObject MemDC, p.handle '
    
    'tranparent blit memdc on usercontrol
    TransBlt UserControl.hDC, x, y, W, H, MemDC, 0, 0, TranslateColor(TOleCol)
    
    DeleteDC MemDC 'its clear, heh
End Sub


Private Function ModifyRect(lpRect As RECT, ByVal Left As Long, ByVal Top As Long, _
               ByVal Right As Long, ByVal Bottom As Long) As RECT
    With ModifyRect
        .Left = lpRect.Left + Left
        .Top = lpRect.Top + Top
        .Right = lpRect.Right + Right
        .Bottom = lpRect.Bottom + Bottom
    End With
End Function
Private Function TranslateColor(ByVal Ole_Color As Long) As Long
        apiTranslateColor Ole_Color, 0, TranslateColor
End Function
Private Function SetRect(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As RECT
  With SetRect
    .Left = Left
    .Top = Top
    .Right = Right
    .Bottom = Bottom
  End With
End Function
Private Sub NormalizeRect(R As RECT)
    Dim C As Long
    If R.Left > R.Right Then
        C = R.Right
        R.Right = R.Left
        R.Left = C
    End If
    If R.Top > R.Bottom Then
        C = R.Top
        R.Top = R.Bottom
        R.Bottom = C
    End If
End Sub
Private Function RoundUp(ByVal num As Single) As Long
    If Int(num) < num Then
        RoundUp = Int(num) + 1
    Else
        RoundUp = num
    End If
End Function
Private Function RectHeight(R As RECT) As Long
    RectHeight = R.Bottom - R.Top
End Function
Private Function RectWidth(R As RECT) As Long
    RectWidth = R.Right - R.Left
End Function
Private Sub DrawText(ByVal hDC As Long, ByVal strText As String, R As RECT, ByVal Format As Long)
    apiDrawText UserControl.hDC, strText, Len(strText), R, Format
End Sub
Private Sub TilePicture(DestRect As RECT, SrcRect As RECT, ByVal SrcDc As Long, Optional UseCliper As Boolean = True, Optional ROp As Long = SRCCOPY)
    
    Dim I As Integer
    Dim J As Integer
    Dim rows As Integer
    Dim ColS As Integer
    Dim destW As Long
    Dim destH As Long
    Dim hDC As Long
    hDC = UserControl.hDC
    
    NormalizeRect DestRect
    NormalizeRect SrcRect
       
    'calculates row and cols
    rows = RoundUp(RectHeight(DestRect) / RectHeight(SrcRect))
    ColS = RoundUp(RectWidth(DestRect) / RectWidth(SrcRect))
    
    destW = RectWidth(SrcRect)
    destH = RectHeight(SrcRect)
   
    'prevents drawing out of specified rectangle
    If UseCliper Then
        SelectClipRgn hDC, ByVal 0
        BeginPath hDC
            With DestRect
                 Rectangle hDC, .Left, .Top, .Right + 1, .Bottom + 1
            End With
        EndPath hDC
        SelectClipPath hDC, RGN_AND
    End If
    
    For I = 0 To rows - 1
        For J = 0 To ColS - 1
            BitBlt hDC, J * destW + DestRect.Left, I * destH + DestRect.Top, destW, destH, SrcDc, _
            SrcRect.Left, SrcRect.Top, ROp
        Next
    Next
    
    If UseCliper Then
        SelectClipRgn hDC, ByVal 0
    End If
End Sub

Private Sub ClearRect(ByVal hDC As Long, lRect As RECT, ByVal Color As Long)
    Dim Brush As Long
    Dim PBrush As Long
    Brush = CreateSolidBrush(Color)
    PBrush = SelectObject(hDC, Brush)
    
    FillRect hDC, lRect, Brush
    DeleteObject SelectObject(hDC, PBrush)
End Sub
'//END GDI####################################
'#############################################

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,3
Public Property Get SizeCW() As Long
Attribute SizeCW.VB_Description = "Corner width."
Attribute SizeCW.VB_ProcData.VB_Invoke_Property = ";Position"
    SizeCW = m_SizeCW
End Property

Public Property Let SizeCW(ByVal New_SizeCW As Long)
        m_SizeCW = New_SizeCW
        PropertyChanged "SizeCW"
        Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,3
Public Property Get SizeCH() As Long
Attribute SizeCH.VB_Description = "Corner height."
Attribute SizeCH.VB_ProcData.VB_Invoke_Property = ";Position"
    SizeCH = m_SizeCH
End Property

Public Property Let SizeCH(ByVal New_SizeCH As Long)
        m_SizeCH = New_SizeCH
        PropertyChanged "SizeCH"
        Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,0,0,0
Public Property Get SkinPicture() As Object
Attribute SkinPicture.VB_Description = "Reference to picture box object."
    Set SkinPicture = m_SkinPicture
End Property

Public Property Set SkinPicture(New_SkinPicture As Object)
    
    
    If (TypeName(New_SkinPicture) <> "PictureBox") And _
       (New_SkinPicture Is Nothing = False) Then
        
        Err.Raise 5, "MyButton::SkinPicture", Err.Description
        Exit Property
    End If
               
    Set m_SkinPicture = New_SkinPicture
    
    If m_SkinPicture Is Nothing = False Then
        m_SkinPictureName = m_SkinPicture.Name
    Else
        m_SkinPictureName = ""
    End If
    
    Refresh
    PropertyChanged "SPN"
End Property
'MemberInfo=13,0,0,
Public Property Get Text() As String
Attribute Text.VB_Description = "Button text."
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    
    
    Refresh
    PropertyChanged "Text"
    
    'setting access key (allows alt + accesskey)
    
    Dim I As Long
    Dim C As String
    
    For I = 1 To Len(New_Text) - 1
        If Mid(New_Text, I, 1) = "&" Then
            C = Mid(New_Text, I + 1, 1)
            If C <> "&" Or C <> " " Then
                UserControl.AccessKeys = C
                PropertyChanged "AccessKey"
                Exit For
            End If
        End If
    Next
   
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get SkinPictureName() As String
Attribute SkinPictureName.VB_Description = "Allows you to set reference at design time."
Attribute SkinPictureName.VB_ProcData.VB_Invoke_Property = ";Appearance"
    'If m_SkinPicture Is Nothing = False Then
        'SkinPictureName = m_SkinPicture.Name
        SkinPictureName = m_SkinPictureName
    'End If
End Property

Public Property Let SkinPictureName(ByVal New_SkinPictureName As String)
    On Error GoTo NotLegalName
    Dim p As Object
    'Debug.Print New_SkinPictureName
    If New_SkinPictureName <> "" Then
        
        Set p = UserControl.Parent.Controls(New_SkinPictureName)
        
        If p Is Nothing = False Then
            Set SkinPicture = p
            'Debug.Print "Setting p"; P.Name
        End If
    Else
        Set m_SkinPicture = Nothing
        'Debug.Print "P is nothing"
        Refresh
    End If
   
    m_SkinPictureName = New_SkinPictureName
    PropertyChanged "SPN"
NotLegalName:
End Property



Private Sub MyTimer_Timer()
    Dim Pos As POINTAPI
    Dim WFP As Long
    
    GetCursorPos Pos
    WFP = WindowFromPoint(Pos.x, Pos.y)
    
    If WFP <> Me.hWnd Then
        UserControl_MouseMove MISION_IMPOSSIBLE, 0, 0, 0
        MyTimer.Enabled = False 'kill that timer at once
    End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case Is = 13
'            RaiseEvent Click
'        Case Is = 27
'            RaiseEvent Click
'        Case Else
'            RaiseEvent Click
'    End Select
    RaiseEvent Click 'optimized heh, highly
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "DisplayAsDefault" Then
'       Debug.Print "Change in " + Me.Text + " - " + CStr(Ambient.DisplayAsDefault)
        DrawButton BTN_NORMAL
    End If
End Sub



Private Sub UserControl_Click()
'
'If Button = vbLeftButton Then
'    If Me.Group = 0 Then Exit Sub
'
'    Dim F As Form
'    Set F = UserControl.Parent
'    Dim C As Control
'    Dim B As MyButton
'
'    For Each C In F.Controls
'        If TypeName(C) = "MyButton" Then
'            Set B = C
'            If B.Group = Me.Group Then
'                If B.Pressed = True Then
'                    B.Pressed = False
'                    Exit For
'                End If
'            End If
'        End If
'    Next
'    Me.Pressed = True
'End If

End Sub

Private Sub UserControl_DblClick()
    If Not m_RightBtn Then
        DrawButton BTN_DOWN
    End If
End Sub

Private Sub UserControl_GotFocus()
    m_HasFocus = True
    If m_BtnDown = False Then DrawButton BTN_FOCUS
End Sub

Private Sub UserControl_Initialize()
'    SkinPictureName = m_SkinPictureName
'    MsgBox "Initialize..." + m_SkinPictureName

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    m_SizeCW = m_def_SizeCW
    m_SizeCH = m_def_SizeCH
    m_Text = Extender.Name
    m_FillWithColor = m_def_FillWithColor
    m_TextColorEnabled = m_def_TextColorEnabled
    m_TextColorDisabled = m_def_TextColorDisabled
    Set UserControl.Font = Ambient.Font
'    m_DisableHover = m_def_DisableHover

    m_DisplaceText = m_def_DisplaceText
    m_DrawFocus = m_def_DrawFocus
    m_TextColorDisabled2 = m_def_TextColorDisabled2
    Set m_Picture = LoadPicture("")
    m_PicturePos = m_def_PicturePos
    m_PictureTColor = m_def_PictureTColor
    m_SkinPictureName = "MyButtonDefSkin"
    m_TextAlign = m_def_TextAlign
    m_DisableHover = m_def_DisableHover
    m_DisableGradient = m_def_DisableGradient
'    m_Group = m_def_Group
    m_Group = m_def_Group
    m_Pressed = m_def_Pressed
    m_OwnerDraw = m_def_OwnerDraw
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    
    If KeyCode = vbKeySpace Then
        m_SpcDown = True
        DrawButton BTN_DOWN
    Else
        m_SpcDown = False
        DrawButton BTN_FOCUS
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    Dim bAlone As Boolean
    If KeyCode = 32 And m_SpcDown And m_State = BTN_DOWN Then
        m_SpcDown = False
        DrawButton BTN_NORMAL
        RaiseEvent Click
        If Me.Group <> 0 Then
            
            Dim F As Form
            Set F = UserControl.Parent
            Dim C As Control
            Dim B As MyButton
            bAlone = True
            For Each C In F.Controls
                If TypeName(C) = CONTROL_NAME Then
                    Set B = C
                    If B.Group = Me.Group And Me.hWnd <> B.hWnd Then
                        bAlone = False
                        If B.Pressed = True Then
                            B.Pressed = False
                            Exit For
                        End If
                    End If
                End If
            Next
            If Not bAlone Then
               Me.Pressed = True
            Else
               Me.Pressed = Not Me.Pressed 'if button is allon in group
            End If
        Else
            DrawButton BTN_FOCUS
        End If
    End If
End Sub

Private Sub UserControl_LostFocus()
    m_HasFocus = False
    DrawButton BTN_NORMAL
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    RaiseEvent MouseDown(Button, Shift, x, y)
    
    If Button = vbRightButton Then
        m_RightBtn = True
    Else
        m_RightBtn = False
    End If
    
    If Button = 1 Then
        m_BtnDown = True
        UserControl_MouseMove Button, Shift, x, y
    End If
   
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_SpcDown Then Exit Sub
    
'   Debug.Print Button
        
     
    'SetCapture hWnd
    If Button = 0 Then
       MyTimer.Enabled = True
    End If
    
    If Button <> MISION_IMPOSSIBLE Then

        
        RaiseEvent MouseMove(Button, Shift, x, y)
        'if pointer is on control
        If m_BtnDown Then
            If m_State <> BTN_DOWN Then
                DrawButton BTN_DOWN
            End If
        Else
            If m_State <> BTN_HOVER Then
                RaiseEvent MouseHover
                DrawButton BTN_HOVER
            End If
            
        End If
    Else
        'if pointer is out of control
        If m_BtnDown Then
            RaiseEvent MouseHover
            DrawButton BTN_HOVER
        Else
            RaiseEvent MouseOut
            If m_HasFocus Then
                DrawButton BTN_FOCUS
            Else
                DrawButton BTN_NORMAL
            End If
            'ReleaseCapture
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    m_BtnDown = False
'    If m_State <> BTN_NORMAL Then
       '' DrawButton BTN_NORMAL
'    End If
    Dim MouseOverControl As Boolean
    
    Dim p As POINTAPI
    GetCursorPos p
    MouseOverControl = (WindowFromPoint(p.x, p.y) = Me.hWnd)
    
    Dim bAlone As Boolean
    'And Me.Pressed = False
    If MouseOverControl Then
    If Button = vbLeftButton And Me.Group <> 0 Then
        
        Dim F As Form
        Set F = UserControl.Parent
        Dim C As Control
        Dim B As MyButton
        
        bAlone = True
        For Each C In F.Controls
        If TypeName(C) = CONTROL_NAME Then
            Set B = C
            If B.Group = Me.Group And Me.hWnd <> B.hWnd Then
                bAlone = False
                If B.Pressed = True Then
                    B.Pressed = False
                    Exit For
                End If
            End If
        End If
        Next
        If Not bAlone Then
            Me.Pressed = True
        Else
            Me.Pressed = Not Me.Pressed 'if button is allon in group
        End If
    End If
    End If

    
    If (Me.Group > 0) And Me.Pressed Then
        DrawButton BTN_DOWN
    ElseIf Button = vbLeftButton Then
        DrawButton BTN_NORMAL
    Else
        If MouseOverControl = False Then DrawButton BTN_NORMAL
    End If

    
    RaiseEvent MouseUp(Button, Shift, x, y)
    
    
    If Button = vbLeftButton Then
        If MouseOverControl Then
            RaiseEvent Click
        Else
            RaiseEvent MouseOut
        End If
        
        If Me.Group = 0 Then
            DrawButton BTN_FOCUS
        Else
            DrawButton BTN_NORMAL
        End If
    End If
End Sub


Private Sub UserControl_Paint()
    Me.Refresh
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_SizeCW = PropBag.ReadProperty("SizeCW", m_def_SizeCW)
    m_SizeCH = PropBag.ReadProperty("SizeCH", m_def_SizeCH)
    m_SkinPictureName = PropBag.ReadProperty("SPN", "")
   
    'Debug.Print "ReadProp SPN:"; m_SkinPictureName
   
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_FillWithColor = PropBag.ReadProperty("FillWithColor", m_def_FillWithColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.AccessKeys = PropBag.ReadProperty("AccessKey", "")
    m_TextColorEnabled = PropBag.ReadProperty("TextColorEnabled", m_def_TextColorEnabled)
    m_TextColorDisabled = PropBag.ReadProperty("TextColorDisabled", m_def_TextColorDisabled)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
'    m_DisableHover = PropBag.ReadProperty("DisableHover", m_def_DisableHover)
'    m_DownTextDX = PropBag.ReadProperty("DownTextDX", m_def_DownTextDX)
'    m_DownTextDY = PropBag.ReadProperty("DownTextDY", m_def_DownTextDY)
    m_DisplaceText = PropBag.ReadProperty("DisplaceText", m_def_DisplaceText)
    m_DrawFocus = PropBag.ReadProperty("DrawFocus", m_def_DrawFocus)
    m_TextColorDisabled2 = PropBag.ReadProperty("TextColorDisabled2", m_def_TextColorDisabled2)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    m_PicturePos = PropBag.ReadProperty("PicturePos", m_def_PicturePos)
    m_PictureTColor = PropBag.ReadProperty("PictureTColor", m_def_PictureTColor)
    m_TextAlign = PropBag.ReadProperty("TextAlign", m_def_TextAlign)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_DisableHover = PropBag.ReadProperty("DisableHover", m_def_DisableHover)
    m_DisableGradient = PropBag.ReadProperty("DisableGradient", m_def_DisableGradient)
'    m_Group = PropBag.ReadProperty("Group", m_def_Group)
    m_Group = PropBag.ReadProperty("Group", m_def_Group)
    m_Pressed = PropBag.ReadProperty("Pressed", m_def_Pressed)
    m_OwnerDraw = PropBag.ReadProperty("OwnerDraw", m_def_OwnerDraw)
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Private Sub UserControl_Show()
    
    SkinPictureName = m_SkinPictureName

   ' Refresh
End Sub

Private Sub UserControl_Terminate()
    Set m_SkinPicture = Nothing
    Set m_Picture = Nothing
    
    'Set UserControl = Nothing
    'Set Me = Nothing
    'Debug.Print "TERMINATE"
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SizeCW", m_SizeCW, m_def_SizeCW)
    Call PropBag.WriteProperty("SizeCH", m_SizeCH, m_def_SizeCH)
    
    'If m_SkinPicture Is Nothing = False Then
        Call PropBag.WriteProperty("SPN", m_SkinPictureName, "")
    'End If
    
    'Debug.Print "Write :"; m_SkinPictureName
    
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("FillWithColor", m_FillWithColor, m_def_FillWithColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("AccessKey", UserControl.AccessKeys, "")
    Call PropBag.WriteProperty("TextColorEnabled", m_TextColorEnabled, m_def_TextColorEnabled)
    Call PropBag.WriteProperty("TextColorDisabled", m_TextColorDisabled, m_def_TextColorDisabled)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)

    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
'    Call PropBag.WriteProperty("DisableHover", m_DisableHover, m_def_DisableHover)
    Call PropBag.WriteProperty("DisplaceText", m_DisplaceText, m_def_DisplaceText)
    Call PropBag.WriteProperty("DrawFocus", m_DrawFocus, m_def_DrawFocus)
    Call PropBag.WriteProperty("TextColorDisabled2", m_TextColorDisabled2, m_def_TextColorDisabled2)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("PicturePos", m_PicturePos, m_def_PicturePos)
    Call PropBag.WriteProperty("PictureTColor", m_PictureTColor, m_def_PictureTColor)
    Call PropBag.WriteProperty("TextAlign", m_TextAlign, m_def_TextAlign)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("DisableHover", m_DisableHover, m_def_DisableHover)
    Call PropBag.WriteProperty("DisableGradient", m_DisableGradient, m_def_DisableGradient)
    Call PropBag.WriteProperty("Group", m_Group, m_def_Group)
    Call PropBag.WriteProperty("Pressed", m_Pressed, m_def_Pressed)
    Call PropBag.WriteProperty("OwnerDraw", m_OwnerDraw, m_def_OwnerDraw)
End Sub


Private Sub DrawButton(ByVal State As Integer)
    If m_DisableHover Then
        If State = BTN_HOVER Then Exit Sub
        'dont draw hover state if m_DisableHover is true
    End If
    If Me.Pressed And Me.Group > 0 Then
        State = BTN_DOWN
    End If
    
    On Error GoTo UnknownError ' element 1 doesnot exists, hihi who cares
     
    If Ambient.DisplayAsDefault Then
        If State = BTN_NORMAL Then
            State = BTN_FOCUS
        End If
    End If
    


    
    Dim PicW As Long
    Dim PicH As Long 'width and height of picture

    Dim PicX As Long
    Dim PicY As Long 'picture pos

    Dim DH As Long  'button height
    Dim DW As Long  'button width
    Dim Align As Long 'text aligment
    Dim bDrawText As Boolean ' if picture is in center text is not drawn
    bDrawText = True

    Align = DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS

    Select Case m_TextAlign
        Case Is = vbLeftJustify:  Align = Align Or DT_LEFT
        Case Is = vbRightJustify: Align = Align Or DT_RIGHT
        Case Is = vbCenter:       Align = Align Or DT_CENTER
    End Select

    DW = UserControl.ScaleWidth
    DH = UserControl.ScaleHeight

    m_State = State
    'if skin picture is not set then just draw text on control
    'If m_SkinPicture Is Nothing Then
        'ClearRect hdc, SetRect(0, 0, dw, DH), TranslateColor(UserControl.BackColor)
        'DrawText hdc, m_Text, SetRect(0, 0, dw, DH), Align
        'If UserControl.AutoRedraw = True Then
        '    UserControl.Refresh
        'End If
        'Exit Sub
    'End If
    
    'UPDATE 1.
    Dim bHasSkin As Boolean
    'if skin picture is not set then it draws standard button
    
    If m_SkinPicture Is Nothing Then
        bHasSkin = False
    Else
        bHasSkin = True
        m_SkinPicture.ScaleMode = vbPixels
    End If

    Dim SrcLeft As Long     'left cordinate of skin in skinpicture
    Dim SrcRight As Long    'right -II-
    Dim FillColor() As Long 'color to fill middle area of button
                            'used if m_FillWithColor is true
    

    Dim H As Long           'height of skinpicture
    Dim W As Long           'width of button skin


    Dim FrameRect As RECT
    ClearRect hDC, SetRect(0, 0, DW, DH), TranslateColor(UserControl.BackColor)
    If bHasSkin = False Then
        'draw button using lines (update 1)
        FrameRect = SetRect(0, 0, DW, DH)
        Select Case State
            Case BTN_DISABLED, BTN_NORMAL, BTN_HOVER
                DrawFrameControl UserControl.hDC, FrameRect, DFC_BUTTON, DFCS_BUTTONPUSH
            Case BTN_DOWN
                DrawFrameControl UserControl.hDC, FrameRect, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_PUSHED
               
            Case BTN_FOCUS
                'default button should be drawn
                '
                DrawFrameControl UserControl.hDC, FrameRect, DFC_BUTTON, DFCS_BUTTONPUSH
        End Select
    Else
        H = m_SkinPicture.ScaleHeight
        W = m_SkinPicture.ScaleWidth / 5
        SrcLeft = (State - 1) * W
        SrcRight = State * W
        
        
        If m_FillWithColor Then
            'get color to fill with from (SrcLeft+m_SizeCW +1 , m_SizeCH+1) on
            'skin picture
            ReDim FillColor(0 To 1)
            FillColor(0) = m_SkinPicture.Point(SrcLeft + m_SizeCW + 1, m_SizeCH + 1)
            FillColor(1) = m_SkinPicture.Point(SrcLeft + m_SizeCW + 1, H - m_SizeCH - 1)
            
            'paint button with fillcolor
            'NOTE: it would be nice if there is gradient file
            'Update 1.
            'with gradient fill looks bette,
            
            'if you wont to disable gradient fill at design time
            'replace if with this one
            
            If FillColor(0) = FillColor(1) Or m_DisableGradient Then
                ClearRect hDC, SetRect(m_SizeCW, m_SizeCH, DW - m_SizeCW, DH - m_SizeCH), FillColor(0)
            Else
                GradientFill hDC, FillColor, m_SizeCW, m_SizeCH, DW - m_SizeCW, DH - m_SizeCH, True
            End If
            'ABOUT ADDING GRADIENT FILL
            'read second color from skin at
            'point (srcleft+cw+1, H -m_sizeCH-1)
            'may be implemented in MyButton2
            '(done in Update 1.)
        Else
            'tile skin 'draw buttton using skin
             TilePicture SetRect(m_SizeCW, m_SizeCH, DW - m_SizeCW, DH - m_SizeCH), _
               SetRect(SrcLeft + m_SizeCW, m_SizeCH, SrcRight - m_SizeCW, H - m_SizeCH), _
               m_SkinPicture.hDC, False, SRCCOPY
        End If
    
        'draws borders
        If (m_SizeCH > 0 And m_SizeCW > 0) Then
            TilePicture SetRect(m_SizeCW, 0, DW, m_SizeCH), _
                SetRect(SrcLeft + m_SizeCW, 0, SrcRight - m_SizeCW, m_SizeCH), _
                m_SkinPicture.hDC, False, SRCCOPY
    
            TilePicture SetRect(m_SizeCW, DH - m_SizeCH, DW, DH), _
                SetRect(SrcLeft + m_SizeCW, H - m_SizeCH, SrcRight - m_SizeCW, H), _
                m_SkinPicture.hDC, False, SRCCOPY
    
            TilePicture SetRect(0, 0, m_SizeCW, DH), _
                SetRect(SrcLeft, m_SizeCH, SrcLeft + m_SizeCW, H - m_SizeCH), _
                m_SkinPicture.hDC, False, SRCCOPY
    
            TilePicture SetRect(DW - m_SizeCW, m_SizeCH, DW, DH - m_SizeCH), _
                SetRect(SrcRight - m_SizeCW, m_SizeCH, SrcRight, H - m_SizeCH), _
                m_SkinPicture.hDC, False, SRCCOPY
    
            'draws corners
            'NOTE: must chage to transparent blit (done)
            TransBlt hDC, 0, 0, m_SizeCW, m_SizeCH, m_SkinPicture.hDC, SrcLeft, 0, &HFF00FF
            TransBlt hDC, 0, DH - m_SizeCH, m_SizeCW, m_SizeCH, m_SkinPicture.hDC, SrcLeft, H - m_SizeCH, &HFF00FF
    
            TransBlt hDC, DW - m_SizeCW, 0, m_SizeCW, m_SizeCH, m_SkinPicture.hDC, SrcRight - m_SizeCW, 0, &HFF00FF
            TransBlt hDC, DW - m_SizeCW, DH - m_SizeCH, m_SizeCW, m_SizeCH, m_SkinPicture.hDC, SrcRight - m_SizeCW, H - m_SizeCH, &HFF00FF
        End If
    End If  'end of if bHasSkin then

'    Dim PColor As Long 'previous color
'
'    PColor = UserControl.ForeColor

    Dim TextRect As RECT
    If State = BTN_DOWN Then
        TextRect = SetRect(m_SizeCW + m_DisplaceText, m_SizeCH + m_DisplaceText, DW - m_SizeCW + m_DisplaceText - 3, DH - m_SizeCH + m_DisplaceText)
    Else
        TextRect = SetRect(m_SizeCW, m_SizeCH, DW - m_SizeCW - 3, DH - m_SizeCH)
    End If
        If m_Picture Is Nothing Then
            If m_State = BTN_DISABLED Then
                'draw text only
                'dont draw text2 if colors are the same
                If m_TextColorDisabled <> m_TextColorDisabled2 Then
                    UserControl.ForeColor = m_TextColorDisabled2
                    TextRect = ModifyRect(TextRect, 1, 1, 1, 1)
                    DrawText hDC, m_Text, TextRect, Align
                    TextRect = ModifyRect(TextRect, -1, -1, -1, -1)
                End If
                UserControl.ForeColor = m_TextColorDisabled
                DrawText hDC, m_Text, TextRect, Align
            Else
                'draw text only
                UserControl.ForeColor = m_TextColorEnabled
                DrawText hDC, m_Text, TextRect, Align
            End If
        Else

            GetBmpSize m_Picture, PicW, PicH
            PicY = (DH - PicH) / 2
            If m_State = BTN_DOWN Then
                PicY = PicY + m_DisplaceText
            End If



            Select Case m_PicturePos
                Case Is = ppLeft
                    PicX = TextRect.Left + 3
                    TextRect.Left = PicX + PicW + TextRect.Left
                Case Is = ppRight
                    PicX = TextRect.Right - PicW - 3 + TextRect.Left - m_SizeCW
                    TextRect.Right = PicX - 3
                Case Is = ppTop
                    PicX = (DW - PicW) / 2 + TextRect.Left - SizeCW
                    PicY = (DH - PicH - 3 - UserControl.TextHeight("I")) / 2 + TextRect.Top - SizeCH
                    TextRect.Top = PicY + PicW + 3
                    TextRect.Bottom = TextRect.Top + UserControl.TextHeight("I") * 1.2
                Case Is = ppBottom
                    TextRect.Top = (DH - PicH - 3 - UserControl.TextHeight("I")) / 2 + TextRect.Top - SizeCH
                    PicX = (DW - PicW) / 2 + TextRect.Left - SizeCW
                    TextRect.Bottom = TextRect.Top + UserControl.TextHeight("I") * 1.2
                    PicY = TextRect.Bottom + 3
                Case Is = ppCenter
                    PicX = (DW - PicW) / 2
                    If BTN_DOWN Then PicX = PicX + m_DisplaceText
                    bDrawText = False
            End Select

'            Debug.Print "State2 "; State

            If m_State = BTN_DISABLED Then
                'draw text and picture disabled
                DrawPictureDisabled m_Picture, PicX, PicY, PicW, PicH
                If m_TextColorDisabled <> m_TextColorDisabled2 Then
                    If bDrawText Then
                        UserControl.ForeColor = m_TextColorDisabled2
                        TextRect = ModifyRect(TextRect, 1, 1, 1, 1)

                        DrawText hDC, m_Text, TextRect, Align
                        TextRect = ModifyRect(TextRect, -1, -1, -1, -1)
                    End If
                End If

                UserControl.ForeColor = m_TextColorDisabled
                If bDrawText Then
                    DrawText hDC, m_Text, TextRect, Align
                End If
            Else
                'draw text and picture enabled
                UserControl.ForeColor = m_TextColorEnabled
                If bDrawText Then
                    DrawText hDC, m_Text, TextRect, Align
                End If
                DrawPicture hDC, m_Picture, PicX, PicY, PicW, PicH, m_PictureTColor
            End If
        End If

    Dim F As Long
    Dim FR As RECT
    F = CLng(m_DrawFocus)
    If F > 0 Then
        If State = BTN_DOWN Or State = BTN_FOCUS Then
            
            FR = SetRect(F, F, DW - F, DH - F)
            UserControl.ForeColor = vbBlack
            SetBkColor hDC, 14215660 'dont touch
            DrawFocusRect hDC, FR
        
        End If
    End If
    
    If Me.OwnerDraw Then
        RaiseEvent OnDrawButton(State)
    End If
    
    If UserControl.AutoRedraw = True Then
        UserControl.Refresh
    End If
Exit Sub
UnknownError:
'Debug.Print Err.Description
'most important line in this function
'i spent about 2 hours to find out
Set m_SkinPicture = Nothing
'removing this line form will not unload properly
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get FillWithColor() As Boolean
Attribute FillWithColor.VB_Description = "Middle area of button is filled with color if true or tiled with skin."
Attribute FillWithColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FillWithColor = m_FillWithColor
End Property

Public Property Let FillWithColor(ByVal New_FillWithColor As Boolean)
    m_FillWithColor = New_FillWithColor
    Refresh
    PropertyChanged "FillWithColor"
End Property


Public Sub Refresh()

    If m_State < 1 Or m_State > 5 Then m_State = 1
    If Enabled Then
        DrawButton m_State
    Else
        DrawButton BTN_DISABLED
    End If
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property
 
Private Function PointInControl(x As Single, y As Single) As Boolean
  If x >= 0 And x <= UserControl.ScaleWidth And _
    y >= 0 And y <= UserControl.ScaleHeight Then
    PointInControl = True
  End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled

    If New_Enabled Then
        DrawButton BTN_NORMAL
    Else
        DrawButton BTN_DISABLED
    End If
    
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TextColorEnabled() As Ole_Color
Attribute TextColorEnabled.VB_Description = "Color of text when its enabled."
Attribute TextColorEnabled.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TextColorEnabled = m_TextColorEnabled
End Property

Public Property Let TextColorEnabled(ByVal New_TextColorEnabled As Ole_Color)
    m_TextColorEnabled = New_TextColorEnabled
    Refresh
    PropertyChanged "TextColorEnabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TextColorDisabled() As Ole_Color
Attribute TextColorDisabled.VB_Description = "Color of text when button is disabled"
Attribute TextColorDisabled.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TextColorDisabled = m_TextColorDisabled
End Property

Public Property Let TextColorDisabled(ByVal New_TextColorDisabled As Ole_Color)
    m_TextColorDisabled = New_TextColorDisabled
    Refresh
    PropertyChanged "TextColorDisabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Refresh
    PropertyChanged "Font"
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = ";Font"
    FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    UserControl.FontUnderline() = New_FontUnderline
    Refresh
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
Attribute FontStrikethru.VB_ProcData.VB_Invoke_Property = ";Font"
    FontStrikethru = UserControl.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    UserControl.FontStrikethru() = New_FontStrikethru
    Refresh
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
Attribute FontSize.VB_ProcData.VB_Invoke_Property = ";Font"
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    Refresh
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
Attribute FontName.VB_ProcData.VB_Invoke_Property = ";Font"
    FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    UserControl.FontName() = New_FontName
    Refresh
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
    FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    UserControl.FontItalic() = New_FontItalic
    Refresh
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
Attribute FontBold.VB_ProcData.VB_Invoke_Property = ";Font"
    FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    UserControl.FontBold() = New_FontBold
    Refresh
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=0,0,0,False
'Public Property Get DisableHover() As Boolean
'    DisableHover = m_DisableHover
'End Property
'
'Public Property Let DisableHover(ByVal New_DisableHover As Boolean)
'    m_DisableHover = New_DisableHover
'    PropertyChanged "DisableHover"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get DisplaceText() As Integer
Attribute DisplaceText.VB_Description = "Displaces text when button is down."
Attribute DisplaceText.VB_ProcData.VB_Invoke_Property = ";Behavior"
    DisplaceText = m_DisplaceText
End Property

Public Property Let DisplaceText(ByVal New_DisplaceText As Integer)
    m_DisplaceText = New_DisplaceText
    PropertyChanged "DisplaceText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get DrawFocus() As Integer
Attribute DrawFocus.VB_Description = "Draws focus."
Attribute DrawFocus.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DrawFocus = m_DrawFocus
End Property

Public Property Let DrawFocus(ByVal New_DrawFocus As Integer)
    m_DrawFocus = New_DrawFocus
    PropertyChanged "DrawFocus"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TextColorDisabled2() As Ole_Color
Attribute TextColorDisabled2.VB_Description = "Color of text when button is disabled that make it looks grayed."
Attribute TextColorDisabled2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TextColorDisabled2 = m_TextColorDisabled2
End Property

Public Property Let TextColorDisabled2(ByVal New_TextColorDisabled2 As Ole_Color)
    m_TextColorDisabled2 = New_TextColorDisabled2
    Refresh
    PropertyChanged "TextColorDisabled2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As StdPicture
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set m_Picture = New_Picture
    Refresh
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get PicturePos() As EnumPicturePos
Attribute PicturePos.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PicturePos = m_PicturePos
End Property

Public Property Let PicturePos(ByVal New_PicturePos As EnumPicturePos)
    m_PicturePos = New_PicturePos
    Refresh
    PropertyChanged "PicturePos"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get PictureTColor() As Ole_Color
Attribute PictureTColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PictureTColor = m_PictureTColor
End Property

Public Property Let PictureTColor(ByVal New_PictureTColor As Ole_Color)
    m_PictureTColor = New_PictureTColor
    Refresh
    PropertyChanged "PictureTColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get TextAlign() As AlignmentConstants
Attribute TextAlign.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TextAlign = m_TextAlign
End Property

Public Property Let TextAlign(ByVal New_TextAlign As AlignmentConstants)
    m_TextAlign = New_TextAlign
    Refresh
    PropertyChanged "TextAlign"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As Ole_Color
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Ole_Color)
    UserControl.BackColor() = New_BackColor
    Refresh
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get DisableHover() As Boolean
    DisableHover = m_DisableHover
End Property

Public Property Let DisableHover(ByVal New_DisableHover As Boolean)
    m_DisableHover = New_DisableHover
    PropertyChanged "DisableHover"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get DisableGradient() As Boolean
Attribute DisableGradient.VB_Description = "Disables gradient fill."
    DisableGradient = m_DisableGradient
End Property

Public Property Let DisableGradient(ByVal New_DisableGradient As Boolean)
    m_DisableGradient = New_DisableGradient
    Refresh
    PropertyChanged "DisableGradient"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Group() As Long
    Group = m_Group
End Property

Public Property Let Group(ByVal New_Group As Long)
    m_Group = New_Group
    PropertyChanged "Group"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Pressed() As Boolean
    Pressed = m_Pressed
End Property

Public Property Let Pressed(ByVal New_Pressed As Boolean)
    m_Pressed = New_Pressed
    PropertyChanged "Pressed"
    DrawButton BTN_NORMAL
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false
Public Property Get OwnerDraw() As Boolean
Attribute OwnerDraw.VB_Description = "Raises OnDrawButton event."
    OwnerDraw = m_OwnerDraw
End Property

Public Property Let OwnerDraw(ByVal New_OwnerDraw As Boolean)
    m_OwnerDraw = New_OwnerDraw
    PropertyChanged "OwnerDraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PaintPicture
Public Sub PaintPicture(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
Attribute PaintPicture.VB_Description = "Draws the contents of a graphics file on a Form, PictureBox, or Printer object."
    UserControl.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
End Sub

