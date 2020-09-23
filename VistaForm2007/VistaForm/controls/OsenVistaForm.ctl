VERSION 5.00
Begin VB.UserControl OsenVistaForm 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2760
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   184
   ToolboxBitmap   =   "OsenVistaForm.ctx":0000
End
Attribute VB_Name = "OsenVistaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''======================================================================================**
''****************************************************************************************
''*  Windows® Vista Emulation Control (OsenVistaForm)                                   **
''*                                                                                     **
''*  Created:     August 15, 2007                                                       **
''*  Updated:     August 16, 2007                                                       **
''*  Purpose:     Form Skinning Control                                                 **
''*  Functions:   (listed)                                                              **
''*  Revision:    ~                                                                     **
''*  Compile:     Native                                                                **
''*  Author:      Osen Kusnadi (http://www.osenxpsuite.net)                             **
''*  Credit:      Paul_Caton@hotmail.com(ucSubclass-self-subclassing :--> cVista Class) **
''*               Carles P.V. (Fast gradient :--> XPaintGradient function)              **
''****************************************************************************************
''* Suggestion and comment are welcome, Please write into support@osenxpsuite.net
''======================================================================================**

Private WithEvents Proc        As cVista
Attribute Proc.VB_VarHelpID = -1

Private PicTop                 As PictureBox
Private PicBottom              As PictureBox
Private PicLeft                As PictureBox
Private PicRight               As PictureBox
Private InitPictObj            As Boolean
Private m_ButtonPos            As Integer
Private m_OldButtonPos         As Integer
Private m_CtlButton            As Integer
Private InUserMode             As Boolean
Private IsInitProc             As Boolean
Private m_pHwnd                As Long
Private m_IsActive             As Boolean
Private CurPOS                 As POINTAPI
Private WorkArea               As RECT
Private m_Icon                 As StdPicture
Private m_Caption              As String

Private Enum eMsg
    ALL_MESSAGES = -1&
    WM_CREATE = &H1&
    WM_DESTROY = &H2&
    WM_MOVE = &H3&
    WM_SIZE = &H5&
    WM_ACTIVATE = &H6&
    WM_SETFOCUS = &H7&
    WM_KILLFOCUS = &H8&
    WM_ENABLE = &HA&
    WM_SETREDRAW = &HB&
    WM_SETTEXT = &HC&
    WM_GETTEXT = &HD&
    WM_GETTEXTLENGTH = &HE&
    WM_PAINT = &HF&
    WM_CLOSE = &H10&
    WM_SHOWWINDOW = &H18&
    WM_GETMINMAXINFO = &H24&
    WM_NCACTIVATE = &H86&
    WM_NCLBUTTONDOWN = &HA1&
    WM_SYSCOMMAND = &H112&
    WM_MOUSEMOVE = &H200&
    WM_LBUTTONDOWN = &H201&
    WM_LBUTTONUP = &H202&
    WM_LBUTTONDBLCLK = &H203&
    WM_MOVING = &H216&
    WM_MOUSELEAVE = &H2A3&
    WM_USER = &H400&
End Enum

#If False Then 'Trick preserves Case of Enums when typing in IDE
    Private ALL_MESSAGES, WM_CREATE, WM_DESTROY, WM_MOVE, WM_SIZE, WM_ACTIVATE, WM_SETFOCUS, WM_KILLFOCUS, WM_ENABLE, WM_SETREDRAW
    Private WM_SETTEXT, WM_GETTEXT, WM_GETTEXTLENGTH, WM_PAINT, WM_CLOSE, WM_SHOWWINDOW, WM_GETMINMAXINFO, WM_NCACTIVATE, WM_NCLBUTTONDOWN
    Private WM_SYSCOMMAND, WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_LBUTTONDBLCLK, WM_MOVING, WM_MOUSELEAVE, WM_USER
#End If

Private Const SC_CLOSE           As Long = &HF060
Private Const SC_MAXIMIZE        As Long = &HF030&
Private Const SC_MINIMIZE        As Long = &HF020&
Private Const SC_RESTORE         As Long = &HF120&
Private Const SC_MOVE            As Long = &HF010&
Private Const HTBOTTOM           As Integer = 15
Private Const HTBOTTOMLEFT       As Integer = 16
Private Const HTBOTTOMRIGHT      As Integer = 17
Private Const HTLEFT             As Integer = 10
Private Const HTRIGHT            As Integer = 11
Private Const HTTOP              As Integer = 12
Private Const HTTOPLEFT          As Integer = 13
Private Const HTTOPRIGHT         As Integer = 14

Private Type OSVERSIONINFO
    dwOSVersionInfoSize             As Long
    dwMajorVersion                  As Long
    dwMinorVersion                  As Long
    dwBuildNumber                   As Long
    dwPlatformId                    As Long
    szCSDVersion                    As String * 128    '  Maintenance string for PSS usage
End Type

Private Type BITMAPINFOHEADER
    biSize                          As Long
    biWidth                         As Long
    biHeight                        As Long
    biPlanes                        As Integer
    biBitCount                      As Integer
    biCompression                   As Long
    biSizeImage                     As Long
    biXPelsPerMeter                 As Long
    biYPelsPerMeter                 As Long
    biClrUsed                       As Long
    biClrImportant                  As Long
End Type

Private Type RECT
    left                            As Long
    top                             As Long
    right                           As Long
    bottom                          As Long
End Type

Private Type POINTAPI
    X                               As Long
    Y                               As Long
End Type

Private Type MINMAXINFO
    ptReserved                      As POINTAPI
    ptMaxSize                       As POINTAPI
    ptMaxPosition                   As POINTAPI
    ptMinTrackSize                  As POINTAPI
    ptMaxTrackSize                  As POINTAPI
End Type

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
    Private TME_HOVER, TME_LEAVE, TME_QUERY, TME_CANCEL
#End If

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                          As Long
    dwFlags                         As TRACKMOUSEEVENT_FLAGS
    hwndTrack                       As Long
    dwHoverTime                     As Long
End Type

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also
'======================================================================================================
Private Sub Proc_WinProcs(pHwnd As Long, uMSG As Long, wParam As Long, lParam As Long)

    Select Case pHwnd

        Case m_pHwnd
            ProcParent uMSG, wParam, lParam

        Case PicTop.hwnd
            ProcTitleBarObject wParam, lParam, uMSG

        Case PicLeft.hwnd
            ProcLeftEgde wParam, uMSG

        Case PicRight.hwnd
            ProcRightEgde wParam, uMSG

        Case PicBottom.hwnd
            ProcBottomEgde wParam, lParam, uMSG
    End Select

End Sub

'======================================================================================================
'PicBottom handler ??? Purpose: Resize Form (Parent Object)
'======================================================================================================
Private Sub ProcBottomEgde(ByVal wParam As Long, ByVal lParam As Long, ByVal uMSG As Long)

    Select Case uMSG

        Case WM_MOUSEMOVE

            If UserControl.Parent.WindowState = 0 Then

                Select Case lParam And &HFFFF&

                    Case Is < 8
                        PicBottom.MousePointer = 6

                    Case Is > PicBottom.ScaleWidth - 8
                        PicBottom.MousePointer = 8

                    Case Else
                        PicBottom.MousePointer = 7
                End Select

            End If

        Case WM_LBUTTONDOWN

            If wParam = 1 Then
               Repos

                Select Case lParam And &HFFFF&

                    Case Is > PicBottom.ScaleWidth - 8
                        ReleaseCapture
                        SendMessage UserControl.Parent.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0&
                       Repos

                    Case Is < 8
                        ReleaseCapture
                        SendMessage UserControl.Parent.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMLEFT, ByVal 0&
                       Repos

                    Case Else
                        ReleaseCapture
                        SendMessage UserControl.Parent.hwnd, WM_NCLBUTTONDOWN, HTBOTTOM, ByVal 0&
                       Repos
                End Select

            End If

    End Select

End Sub

'======================================================================================================
'PicLeft handler ??? Purpose: Resize Form (Parent Object)
'======================================================================================================
Private Sub ProcLeftEgde(ByVal wParam As Long, ByVal uMSG As Long)

    If UserControl.Parent.WindowState = 0 Then

        Select Case uMSG

            Case WM_MOUSEMOVE
                PicLeft.MousePointer = 9

            Case WM_LBUTTONDOWN

                If wParam = 1 Then
                   Repos
                    ReleaseCapture
                    SendMessage m_pHwnd, WM_NCLBUTTONDOWN, HTLEFT, 0&
                End If

        End Select

    End If

End Sub

'======================================================================================================
'Form handler ??? Purpose: ...
'======================================================================================================
Private Sub ProcParent(ByVal uMSG As Long, ByVal wParam As Long, lParam As Long)
    Dim udtMINMAXINFO As MINMAXINFO
    Dim nWidthPixels  As Long
    Dim nHeightPixels As Long

    Select Case uMSG

        Case WM_ACTIVATE
            m_IsActive = wParam
            Repos

        Case WM_SYSCOMMAND

            Select Case wParam

                Case SC_CLOSE      ''debug.print "Kapandim..."
                    Proc.DetachMSG
                    Unload UserControl.Parent

                Case SC_RESTORE    ''debug.print "Küçüldüm..."
                    UserControl.Parent.WindowState = 0
                   Repos

                Case SC_MAXIMIZE   ''debug.print "Büyüdüm..."
                    UserControl.Parent.WindowState = 2
                   Repos

                Case SC_MOVE       ''debug.print "Taþýnýyorum..."

                    If UserControl.Parent.WindowState = 2 Then
                        Exit Sub
                    End If

                    ReleaseCapture
                    SendMessage m_pHwnd, WM_SYSCOMMAND, &HF012&, 0&

                Case SC_MINIMIZE   ''debug.print "Minimize Oluyom..."

                Case Else
                   Repos
            End Select

        Case WM_SIZE, WM_MOVING, WM_LBUTTONDOWN
           Repos

        Case WM_NCACTIVATE
            m_IsActive = wParam
            Repos

        Case WM_SETFOCUS
           Repos

        Case WM_GETMINMAXINFO
            GetWorkAreA
            nWidthPixels = WorkArea.right
            nHeightPixels = WorkArea.bottom
            CopyMemory udtMINMAXINFO, ByVal lParam, 40&

            With udtMINMAXINFO
                .ptMaxSize.X = nWidthPixels '- (nWidthPixels \ 4)
                .ptMaxSize.Y = nHeightPixels '- 30 '(nHeightPixels \ 4)
                .ptMaxPosition.X = 0 'nWidthPixels \ 8
                .ptMaxPosition.Y = 0 'nHeightPixels \ 8
            End With 'UDTMINMAXINFO

            CopyMemory ByVal lParam, udtMINMAXINFO, 40&

        Case WM_SHOWWINDOW
            m_IsActive = True
           Repos
            UserControl.Parent.Controls(Ambient.DisplayName).ZOrder
    End Select

End Sub

'======================================================================================================
'PicRight handler ??? Purpose: Resize Form (Parent Object)
'======================================================================================================
Private Sub ProcRightEgde(ByVal wParam As Long, ByVal uMSG As Long)

    If UserControl.Parent.WindowState = 0 Then

        Select Case uMSG

            Case WM_MOUSEMOVE
                PicRight.MousePointer = 9

            Case WM_LBUTTONDOWN

                If wParam = 1 Then
                   Repos
                    ReleaseCapture
                    SendMessage m_pHwnd, WM_NCLBUTTONDOWN, HTRIGHT, 0&
                End If

        End Select

    End If

End Sub

'======================================================================================================
'PicTop handler ??? Purpose: Resize Form (Parent Object)
'======================================================================================================
Private Sub ProcTitleBarObject(wParam As Long, lParam As Long, ByVal uMSG As Long)
    Dim tme As TRACKMOUSEEVENT_STRUCT

    Select Case uMSG

        Case WM_MOUSELEAVE
            m_ButtonPos = 0

            If m_OldButtonPos <> m_ButtonPos Then
                If m_IsActive Then
                    If UserControl.Parent.WindowState = 2 Then
                        m_CtlButton = 181
                    Else 'NOT USERCONTROL.PARENT.WINDOWSTATE...
                        m_CtlButton = 191
                    End If

                Else 'M_ISACTIVE = FALSE/0
                    m_CtlButton = 171
                End If

                DrawTitle m_CtlButton
                m_OldButtonPos = m_ButtonPos
            End If

        Case WM_LBUTTONDBLCLK

            If UserControl.Parent.WindowState = 0 Then
                SendMessage m_pHwnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0&
            Else 'NOT USERCONTROL.PARENT.WINDOWSTATE...
                SendMessage m_pHwnd, WM_SYSCOMMAND, SC_RESTORE, 0&
            End If

        Case WM_MOUSEMOVE
            CurPOS = GetXY(lParam)
            DefineButton CurPOS

            If wParam = 1 And CurPOS.X < PicTop.ScaleWidth - 100 Then
                ReleaseCapture
                SendMessage m_pHwnd, &HA1, 2, 0
            Else 'NOT WPARAM...

                If CurPOS.Y < 18 Then
                    If CurPOS.X > PicTop.ScaleWidth - 101 And CurPOS.X < PicTop.ScaleWidth - 8 Then
                        PicTop.MousePointer = 0

                        With tme
                            .cbSize = Len(tme)
                            .dwFlags = TME_LEAVE
                            .hwndTrack = PicTop.hwnd
                        End With 'TME

                        If m_OldButtonPos <> m_ButtonPos Then
                            m_CtlButton = GetResButtonId(wParam)
                            DrawTitle m_CtlButton
                            m_OldButtonPos = m_ButtonPos
                        End If

                        TrackMouseEvent tme
                    Else 'NOT CURPOS.X...

                        If CurPOS.Y < 5 Then

                            Select Case CurPOS.X

                                Case Is < 8
                                    PicTop.MousePointer = 8

                                Case Is > PicTop.ScaleWidth - 8
                                    PicTop.MousePointer = 6

                                Case PicTop.ScaleWidth - 100 To PicTop.ScaleWidth - 8
                                    PicTop.MousePointer = 0

                                Case Else
                                    PicTop.MousePointer = 7
                            End Select

                        End If
                    End If

                Else 'NOT CURPOS.Y...
                    PicTop.MousePointer = 0
                End If
            End If

        Case WM_LBUTTONDOWN
            CurPOS = GetXY(lParam)

            If CurPOS.Y < 5 Then

                Select Case CurPOS.X

                    Case Is < 8
                        ReleaseCapture
                        SendMessage m_pHwnd, WM_NCLBUTTONDOWN, HTTOPLEFT, ByVal 0&
                       Repos

                    Case Is > PicTop.ScaleWidth - 8
                        ReleaseCapture
                        SendMessage m_pHwnd, WM_NCLBUTTONDOWN, HTTOPRIGHT, ByVal 0&
                       Repos

                    Case Else
                        ReleaseCapture
                        SendMessage m_pHwnd, WM_NCLBUTTONDOWN, HTTOP, ByVal 0&
                       Repos
                End Select

            Else 'NOT CURPOS.Y...
                PicTop.MousePointer = 0

                If CurPOS.X > PicTop.ScaleWidth - 100 Then
                    If CurPOS.Y < 17 Then
                        DefineButton CurPOS

                        If m_OldButtonPos <> m_CtlButton Then
                            m_CtlButton = GetResButtonId(wParam)
                            DrawTitle m_CtlButton
                            m_OldButtonPos = m_CtlButton
                        End If
                    End If
                End If
            End If

        Case WM_LBUTTONUP
            CurPOS = GetXY(lParam)

            If CurPOS.X > PicTop.ScaleWidth - 100 Then
                If CurPOS.Y < 17 Then
                    DefineButton CurPOS

                    If m_OldButtonPos <> m_CtlButton Then
                        m_CtlButton = GetResButtonId(wParam)
                        DrawTitle m_CtlButton
                        m_OldButtonPos = m_CtlButton
                    End If
                End If

                Select Case m_ButtonPos

                    Case 3
                        SendMessage m_pHwnd, WM_SYSCOMMAND, SC_CLOSE, 0&

                    Case 2

                        If UserControl.Parent.WindowState = 0 Then
                            SendMessage m_pHwnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0&
                        Else 'NOT USERCONTROL.PARENT.WINDOWSTATE...
                            SendMessage m_pHwnd, WM_SYSCOMMAND, SC_RESTORE, 0&
                        End If

                    Case 1

                        SendMessage m_pHwnd, WM_SYSCOMMAND, SC_MINIMIZE, 0&

                End Select

            End If

    End Select

End Sub

'======================================================================================================
' Draw Caption on PicTop Object ... (TitleBar)
'======================================================================================================
Private Function DrawTextXP(ByVal lngHdc As Long, ByVal sCaption As String, ByRef rcText As RECT, Optional Flags As Long) As Long

    If Len(sCaption) Then
        If IsNT Then
            DrawTextXP = DrawTextW(lngHdc, StrPtr(sCaption), -1, rcText, Flags Or &H8000&)
        Else 'ISNT = FALSE/0
            DrawTextXP = DrawText(lngHdc, sCaption, -1, rcText, Flags Or &H8000&)
        End If
    End If

End Function

'======================================================================================================
' Convert lParam value (from subclass handler) into X,Y coordinate (Cursor position)
'======================================================================================================
Private Function GetXY(ByVal lParam As Long) As POINTAPI
    GetXY.X = lParam And &HFFFF&
    GetXY.Y = lParam \ &H10000 And &HFFFF&
End Function

'======================================================================================================
' Determine This Project running on WinNT or not...
'======================================================================================================
Private Property Get IsNT() As Boolean
    Dim lPlatform As Long
    Dim uVer      As OSVERSIONINFO
    uVer.dwOSVersionInfoSize = Len(uVer)

    If GetVersionEx(uVer) Then
        lPlatform = uVer.dwPlatformId
    End If

    IsNT = (lPlatform = 2)
End Property

'================================================
' Module:        mGradient.bas
' Author:        Carles P.V. - 2005
' Dependencies:  None
' Last revision: 2005.05.13
'================================================
Private Sub XPaintGradient(ByVal lngHdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lngWidth As Long, ByVal lngHeight As Long, Optional ByVal Color1 As Long, Optional ByVal Color2 As Long = vbWhite, Optional ByVal GradientDirection As Long = 1)
    Dim uBIH    As BITMAPINFOHEADER
    Dim lBits() As Long
    Dim lGrad() As Long
    Dim R1      As Long
    Dim G1      As Long
    Dim b1      As Long
    Dim R2      As Long
    Dim G2      As Long
    Dim b2      As Long
    Dim dR      As Long
    Dim dG      As Long
    Dim dB      As Long
    Dim Scan    As Long
    Dim i       As Long
    Dim iEnd    As Long
    Dim iOffset As Long
    Dim j       As Long
    Dim jEnd    As Long
    Dim iGrad   As Long
    On Error Resume Next

    If Not (lngWidth < 1 Or lngHeight < 1) Then
        Color1 = Color1 And &HFFFFFF
        R1 = Color1 Mod &H100&
        Color1 = Color1 \ &H100&
        G1 = Color1 Mod &H100&
        Color1 = Color1 \ &H100&
        b1 = Color1 Mod &H100&
        Color2 = Color2 And &HFFFFFF
        R2 = Color2 Mod &H100&
        Color2 = Color2 \ &H100&
        G2 = Color2 Mod &H100&
        Color2 = Color2 \ &H100&
        b2 = Color2 Mod &H100&
        dR = R2 - R1
        dG = G2 - G1
        dB = b2 - b1

        Select Case GradientDirection

            Case 0
                ReDim lGrad(0 To lngWidth - 1) As Long

            Case 1
                ReDim lGrad(0 To lngHeight - 1) As Long

            Case Else
                ReDim lGrad(0 To lngWidth + lngHeight - 2) As Long
        End Select

        iEnd = UBound(lGrad())

        If iEnd = 0 Then
            lGrad(0) = (b1 \ 2 + b2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
        Else 'NOT IEND...

            For i = 0 To iEnd
                lGrad(i) = b1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
            Next i

        End If

        ReDim lBits(lngWidth * lngHeight - 1) As Long
        iEnd = lngWidth - 1
        jEnd = lngHeight - 1
        Scan = lngWidth

        Select Case GradientDirection

            Case 0

                For j = 0 To jEnd
                    For i = iOffset To iEnd + iOffset
                        lBits(i) = lGrad(i - iOffset)
                    Next i

                    iOffset = iOffset + Scan
                Next j

            Case 1

                For j = jEnd To 0 Step -1
                    For i = iOffset To iEnd + iOffset
                        lBits(i) = lGrad(j)
                    Next i

                    iOffset = iOffset + Scan
                Next j

            Case 2
                iOffset = jEnd * Scan

                For j = 1 To jEnd + 1
                    For i = iOffset To iEnd + iOffset
                        lBits(i) = lGrad(iGrad)
                        iGrad = iGrad + 1
                    Next i

                    iOffset = iOffset - Scan
                    iGrad = j
                Next j

            Case Else
                iOffset = 0

                For j = 1 To jEnd + 1
                    For i = iOffset To iEnd + iOffset
                        lBits(i) = lGrad(iGrad)
                        iGrad = iGrad + 1
                    Next i

                    iOffset = iOffset + Scan
                    iGrad = j
                Next j

        End Select

        With uBIH
            .biSize = 40
            .biPlanes = 1
            .biBitCount = 32
            .biWidth = lngWidth
            .biHeight = lngHeight
        End With 'UBIH

        StretchDIBits lngHdc, X, Y, lngWidth, lngHeight, 0, 0, lngWidth, lngHeight, lBits(0), uBIH, 0, vbSrcCopy
        On Error GoTo 0
    End If

End Sub


Public Property Get Caption() As String
    Caption = m_Caption
End Property

' Change Caption Property ...
Public Property Let Caption(ByVal New_Caption As String)

    m_Caption = New_Caption
    PropertyChanged "Caption"
    
    ' Redraw
    Repos
End Property

'======================================================================================================
' Check cursor position to determine Button Position (MinButton=1,MaxButon=2,CloseButton=3, none: 0)
'======================================================================================================
Private Sub DefineButton(CurPOS As POINTAPI)
    Dim X As Long
    m_ButtonPos = 0

    If CurPOS.X > PicTop.ScaleWidth - 100 Then
        If CurPOS.Y < 17 Then
            X = PicTop.ScaleWidth - 100

            Select Case CurPOS.X

                Case Is < X + 25
                    m_ButtonPos = 1

                Case Is < X + 50
                    m_ButtonPos = 2

                Case Is < X + 93
                    m_ButtonPos = 3

                Case Else
                    m_ButtonPos = 0
            End Select

        End If
    End If

End Sub

'======================================================================================================
' Draw Caption at the TitleBar (PicTop Object)
'======================================================================================================
Private Sub DrawCaption()
    Dim mRect      As RECT
    Dim m_CapColor As Long
    On Error Resume Next

    With mRect
        .left = IIf((Not m_Icon Is Nothing), 28, 7)
        .top = 6
        .bottom = 28
        .right = PicTop.ScaleWidth - 100
    End With 'MRECT

    If m_IsActive Then
        SetTextColor PicTop.hdc, RGB(10, 24, 131)
        m_CapColor = vbWhite
    Else 'ISACTIVE = FALSE/0'M_ISACTIVE = FALSE/0
        SetTextColor PicTop.hdc, &HE0E0E0
        m_CapColor = vbWhite
    End If

    If m_IsActive Then
        mRect.left = mRect.left + 1
        mRect.top = mRect.top + 1
        DrawTextXP PicTop.hdc, m_Caption, mRect, &H0
    End If

    SetTextColor PicTop.hdc, m_CapColor
    mRect.left = mRect.left - Abs(m_IsActive)
    mRect.top = mRect.top - Abs(m_IsActive)
    DrawTextXP PicTop.hdc, m_Caption, mRect, &H0
    On Error GoTo 0
End Sub

'======================================================================================================
' Drawing TitleBar ... (Skin, Caption, and an Icon if available)...
'======================================================================================================
Private Sub DrawTitle(Optional ButtonID As Integer = 0)
    Dim k As Integer
    On Error Resume Next

    If m_IsActive Then
        k = 1
    Else 'M_ISACTIVE = FALSE/0
        k = 0
    End If

    With PicTop ' titlebar
        .AutoRedraw = True
        .Cls
        .PaintPicture LoadResPicture(110 + k, 0), 0, 0, 10, 28
        .PaintPicture LoadResPicture(110 + k, 0), 10, 0, .ScaleWidth, 28, 9, 0, 1, 28

        If ButtonID = 0 Then
            If m_IsActive Then
                If UserControl.Parent.WindowState = 2 Then
                    ButtonID = 181
                Else 'NOT USERCONTROL.PARENT.WINDOWSTATE...
                    ButtonID = 191
                End If

            Else 'M_ISACTIVE = FALSE/0
                ButtonID = 171
            End If
        End If

        .PaintPicture LoadResPicture(ButtonID, 0), .ScaleWidth - 108, 0, 108, 28

        If (Not m_Icon Is Nothing) Then
            .PaintPicture m_Icon, 6, 6, 16, 16
        End If

        DrawCaption
        .Refresh
        .AutoRedraw = False
    End With 'PICTOP

    On Error GoTo 0
End Sub

'======================================================================================================
' Draw Form skinning control ...
'======================================================================================================
Private Sub DrawVista()
    On Error Resume Next
    Dim l As Integer
    
    ' Draw TitleBar ...
    DrawTitle

    ' Check this parent (form object) is active or deactive
    If m_IsActive Then
        l = 1
    Else 'M_ISACTIVE = FALSE/0
        l = 0
    End If

    ' drawskin on PicBottom Object
    With PicBottom ' bottom
        .AutoRedraw = True
        .Cls
        ' Draw BottomLeft Picture
        .PaintPicture LoadResPicture(130 + l, 0), 0, 0
        
        ' Draw BottomMiddle Picture
        .PaintPicture LoadResPicture(140 + l, 0), 8, 0, .ScaleWidth, 8, 0, 0, 1, 8
        
        ' Draw BottomRight Picture
        .PaintPicture LoadResPicture(150 + l, 0), .ScaleWidth - 8, 0
        
        ' Refresh object ...
        .Refresh
        .AutoRedraw = False
    End With 'PICBOTTOM

    With PicLeft ' left
        .AutoRedraw = True
        
        ' clean up
        .Cls
        
        ' Draw Left Side Picture on Picleft ...
        .PaintPicture LoadResPicture(120 + l, 0), 0, 0, 8, .ScaleHeight
        
        ' Draw Gradient color, 25% from scaleheight of the PicLeft Object
        If m_IsActive Then
            XPaintGradient .hdc, 2, 0, 4, (.ScaleHeight * 0.25), 8750469, 14606046
        Else '12369084
            XPaintGradient .hdc, 2, 0, 4, (.ScaleHeight * 0.25), 12369084, 14606046
        End If
        
        ' Refresh Object...
        .Refresh
        .AutoRedraw = False
    End With 'PICLEFT

    With PicRight ' right
        .AutoRedraw = True
        
        ' Clean Up
        .Cls
        
        ' Draw Right side picture on PicRight ...
        .PaintPicture LoadResPicture(160 + l, 0), 0, 0, 8, .ScaleHeight
        
        ' Draw Gradient color, 25% from scaleheight of the PicRight Object
        If m_IsActive Then
            XPaintGradient .hdc, 2, 0, 4, (.ScaleHeight * 0.25), 8750469, 14606046
        Else '12369084
            XPaintGradient .hdc, 2, 0, 4, (.ScaleHeight * 0.25), 12369084, 14606046
        End If
        
        ' Refresh
        .Refresh
        .AutoRedraw = False
    End With 'PICRIGHT

    On Error GoTo 0
End Sub

Private Function GetResButtonId(ByVal wParam As Long) As Integer
    On Error Resume Next

    If m_IsActive Then
        If UserControl.Parent.WindowState = 2 Then
            If wParam And m_ButtonPos > 0 Then
                GetResButtonId = 181 + m_ButtonPos + 3
            Else 'NOT WPARAM...
                GetResButtonId = 181 + m_ButtonPos
            End If

        Else 'NOT USERCONTROL.PARENT.WINDOWSTATE...

            If wParam And m_ButtonPos > 0 Then
                GetResButtonId = 191 + m_ButtonPos + 3
            Else 'NOT WPARAM...
                GetResButtonId = 191 + m_ButtonPos
            End If
        End If

    Else 'M_ISACTIVE = FALSE/0

        If UserControl.Parent.WindowState = 2 Then
            GetResButtonId = 171 + m_ButtonPos + 3
        Else 'NOT USERCONTROL.PARENT.WINDOWSTATE...
            GetResButtonId = 171 + m_ButtonPos
        End If
    End If

    On Error GoTo 0
End Function

Private Sub GetWorkAreA()
    SystemParametersInfo 48&, 0&, WorkArea, 0&
End Sub

Public Property Get Icon() As StdPicture
    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal IpValue As StdPicture)
    Set m_Icon = IpValue
    PropertyChanged "icon"

    If Not m_Icon Is Nothing Then
        Set UserControl.Parent.Icon = m_Icon
    End If
    
    ' Redraw skin ...
    Repos
End Property

Private Sub InitControl()

    If Not InitPictObj Then
        InitPictObj = True

        ' Add new pictureBox object into usercontrol at RunTime...
        With UserControl
            Set PicTop = .Controls.Add("VB.PictureBox", "PicTop", Me)
            Set PicBottom = .Controls.Add("VB.PictureBox", "PicBottom", Me)
            Set PicLeft = .Controls.Add("VB.PictureBox", "PicLeft", Me)
            Set PicRight = .Controls.Add("VB.PictureBox", "PicRight", Me)
        End With 'UserControl

        ' Set Default PictureBox Style
        SetPicStyle PicTop
        SetPicStyle PicBottom
        SetPicStyle PicLeft
        SetPicStyle PicRight

        With PicTop.Font
            .Bold = True
            .Size = 10
        End With 'PICTOP.FONT

    End If

End Sub

' Start Subclass
Private Sub InitSubClass()

    If InUserMode Then
    
        ' Get handle window from parent object
        m_pHwnd = UserControl.Parent.hwnd
        
        ' create new instance of cVista class object
        Set Proc = New cVista
        
        IsInitProc = True

        With Proc
        
            ' Parent object handle (form object)
            .Start m_pHwnd
            .AttachAfterMSG WM_SYSCOMMAND
            .AttachAfterMSG WM_MOVING
            .AttachAfterMSG WM_SIZE
            .AttachAfterMSG WM_SHOWWINDOW
            .AttachAfterMSG WM_SETFOCUS
            .AttachAfterMSG WM_NCLBUTTONDOWN
            .AttachAfterMSG WM_LBUTTONDOWN
            .AttachAfterMSG WM_PAINT
            .AttachAfterMSG WM_NCACTIVATE
            .AttachMSG WM_GETMINMAXINFO
            
            ' Start subclass for PicTop object
            .Start PicTop.hwnd
            .AttachAfterMSG WM_LBUTTONDOWN
            .AttachAfterMSG WM_MOUSEMOVE
            .AttachAfterMSG WM_LBUTTONDBLCLK
            .AttachAfterMSG WM_LBUTTONUP
            .AttachAfterMSG WM_MOUSELEAVE
            .AttachAfterMSG WM_NCLBUTTONDOWN
            
            ' Start subclass for PicLeft object
            .Start PicLeft.hwnd
            .AttachAfterMSG WM_LBUTTONDOWN
            .AttachAfterMSG WM_MOUSEMOVE
            
            ' Start subclass for PicRight object
            .Start PicRight.hwnd
            .AttachAfterMSG WM_LBUTTONDOWN
            .AttachAfterMSG WM_MOUSEMOVE
            
            ' Start subclass for PicBottom object
            .Start PicBottom.hwnd
            .AttachAfterMSG WM_LBUTTONDOWN
            .AttachAfterMSG WM_MOUSEMOVE
            
        End With 'PROC
        
        ' Redraw skin and re-position...
        Repos
        
    End If

End Sub

' Make roundegde at the parent object
Private Sub MakeRegion()

    Dim rgn1    As Long
    Dim rgn2    As Long
    Dim rgnNorm As Long
    Dim hResult As Long
    Dim Wi      As Long
    Dim He      As Long
    
    Wi = UserControl.Parent.ScaleWidth
    He = UserControl.Parent.ScaleHeight
    
    ' top left
    rgnNorm = CreateRectRgn(0, 0, Wi, He)
    rgn2 = CreateRectRgn(0, 0, 0, 0)
    rgn1 = CreateRectRgn(0, 0, 2, 2)
    CombineRgn rgn2, rgnNorm, rgn1, 4
    DeleteObject rgn1
    rgn1 = CreateRectRgn(0, 0, 1, 4)
    CombineRgn rgn2, rgn2, rgn1, 4
    DeleteObject rgn1
    rgn1 = CreateRectRgn(0, 0, 4, 1)
    CombineRgn rgn2, rgn2, rgn1, 4
    DeleteObject rgn1
    
    ' Bottom Left
    rgn1 = CreateRectRgn(0, He, 2, He - 2)
    CombineRgn rgnNorm, rgn2, rgn1, 4
    DeleteObject rgn1
    rgn1 = CreateRectRgn(0, He, 1, He - 4)
    CombineRgn rgnNorm, rgnNorm, rgn1, 4
    DeleteObject rgn1
    rgn1 = CreateRectRgn(0, He, 4, He - 1)
    CombineRgn rgnNorm, rgnNorm, rgn1, 4
    DeleteObject rgn1
    
    ' Top Right
    rgn1 = CreateRectRgn(Wi, 0, Wi - 2, 2)
    CombineRgn rgn2, rgnNorm, rgn1, 4
    DeleteObject rgn1
    rgn1 = CreateRectRgn(Wi, 0, Wi - 4, 1)
    CombineRgn rgn2, rgn2, rgn1, 4
    DeleteObject rgn1
    rgn1 = CreateRectRgn(Wi, 0, Wi - 1, 4)
    CombineRgn rgn2, rgn2, rgn1, 4
    DeleteObject rgn1
    
    ' Bottom Right
    rgn1 = CreateRectRgn(Wi, He, Wi - 2, He - 2)
    CombineRgn rgnNorm, rgn2, rgn1, 4
    DeleteObject rgn1
    rgn1 = CreateRectRgn(Wi, He, Wi - 1, He - 4)
    CombineRgn rgnNorm, rgnNorm, rgn1, 4
    DeleteObject rgn1
    rgn1 = CreateRectRgn(Wi, He, Wi - 4, He - 1)
    CombineRgn rgnNorm, rgnNorm, rgn1, 4
    hResult = SetWindowRgn(UserControl.Parent.hwnd, rgnNorm, True)
    
    ' Clean Up
    DeleteObject rgn1
    DeleteObject rgn2
    DeleteObject rgnNorm
    DeleteObject hResult
    
End Sub


' Reposition skinning object and drawing skin
Private Sub Repos()

    On Error Resume Next

    If InUserMode Then

        With UserControl

            If .Parent.WindowState <> 1 Then
            
                .Height = .Parent.Height

                If .Height > 36 Then
                    PicTop.Move 0, 0, .ScaleWidth, 28
                    PicLeft.Move 0, 28, 8, .ScaleHeight - 36
                    PicRight.Move .ScaleWidth - 8, 28, 8, .ScaleHeight - 36
                    PicBottom.Move 0, .ScaleHeight - 8, .ScaleWidth, 8
                End If
                
                MakeRegion
                
                ' Draw TitleBar ...
                DrawVista
                
            End If

        End With 'UserControl
    Else
        ' Draw TitleBar ...
        DrawVista
    End If

    On Error GoTo 0
End Sub

' Set standard pictureBox style...
Private Sub SetPicStyle(IpPic As PictureBox)

    With IpPic
        .Appearance = 0
        .BorderStyle = 0
        .AutoRedraw = False
        .BackColor = vbWhite
        .ScaleMode = 3
        .Visible = True
    End With 'IPPIC

End Sub

Private Sub UserControl_Initialize()
    m_IsActive = True
    InitControl
End Sub

Private Sub UserControl_InitProperties()

    With UserControl.Parent
        .BorderStyle = 0
        .ShowInTaskbar = True
        .Controls(Ambient.DisplayName).Align = 1
        .ScaleMode = 3
    End With 'USERCONTROL.PARENT
    Set m_Icon = UserControl.Parent.Icon
    m_Caption = "VistaForm"
    InitControl
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    InUserMode = Ambient.UserMode
    m_Caption = PropBag.ReadProperty("Caption", "VistaForm")
    Set m_Icon = PropBag.ReadProperty("icon", Nothing)
    m_IsActive = True
    InitControl
    InitSubClass
    DrawVista
    
End Sub

Private Sub UserControl_Resize()

    If Not InUserMode Then
        UserControl.Height = 420
        PicTop.Width = UserControl.ScaleWidth
        DrawVista
    End If

End Sub

Private Sub UserControl_Terminate()

    If IsInitProc Then
        Proc.DetachMSG
        Set Proc = Nothing
        IsInitProc = False
    End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", m_Caption, "VistaForm"
    PropBag.WriteProperty "icon", m_Icon, Nothing
End Sub
