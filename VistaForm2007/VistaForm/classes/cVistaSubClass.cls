VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 28513
Attribute VB_Description = "OSENXPSUITE2006.CLS_MySubclass"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'==================================================================================================
' ucSubclass - A sample UserControl demonstrating self-subclassing
'
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
' v1.0.0000 20040525 First cut.....................................................................
' v1.1.0000 20040602 Multi-subclassing version.....................................................
' v1.1.0001 20040604 Optimized the subclass code...................................................
' v1.1.0002 20040607 Substituted byte arrays for strings for the code buffers......................
' v1.1.0003 20040618 Re-patch when adding extra hWnds..............................................
' v1.1.0004 20040619 Optimized to death version....................................................
' v1.1.0005 20040620 Use allocated memory for code buffers, no need to re-patch....................
' v1.1.0006 20040628 Better protection in zIdx, improved comments..................................
' v1.1.0007 20040629 Fixed InIDE patching oops.....................................................
' v1.1.0008 20040910 Fixed bug in UserControl_Terminate, zSubclass_Proc procedure hidden...........

Private Type tSubData
    hwnd                              As Long        'Handle of the window being subclassed
    nAddrSub                          As Long
    nAddrOrig                         As Long        'The address of the pre-existing WndProc
    nMsgCntA                          As Long        'Msg after table entry count
    nMsgCntB                          As Long        'Msg before table entry count
    aMsgTblA()                        As Long        'Msg after table array
    aMsgTblB()                        As Long        'Msg Before table array
End Type

Private Enum eMsgWhen
    MSG_AFTER = 1
    MSG_BEFORE = 2
    MSG_BEFORE_AND_AFTER = 3
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE

    Private MSG_AFTER, MSG_BEFORE, MSG_BEFORE_AND_AFTER
#End If

Private Const ALL_MESSAGES        As Long = -1 'All messages added or deleted

Private Const CODE_LEN            As Long = 197 'Length of the machine code in bytes

Private Const GWL_WNDPROC         As Long = -4

Private Const PATCH_04            As Long = 88 'Table B (before) address patch offset

Private Const PATCH_05            As Long = 93 'Table B (before) entry count patch offset

Private Const PATCH_08            As Long = 132 'Table A (after) address patch offset

Private Const PATCH_09            As Long = 137 'Table A (after) entry count patch offset

Private sc_aBuf(1 To CODE_LEN)    As Byte        'Code buffer byte array

Private sc_pCWP                   As Long        'Address of the CallWindowsProc

Private sc_pEbMode                As Long

Private sc_pSWL                   As Long        'Address of the SetWindowsLong function

Private sc_aSubData()             As tSubData    '//-- Array of Subclass Values

Private m_Allowinide              As Boolean

Private m_Status                  As Boolean

Private m_LastHWND                As Long

Public Event WinProcs(pHwnd As Long, uMSG As Long, wParam As Long, lParam As Long)

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMSG As Long, ByRef wParam As Long, ByRef lParam As Long)
    RaiseEvent WinProcs(lng_hWnd, uMSG, wParam, lParam)
End Sub

Public Function Start(ByVal lng_hWnd As Long) As Long
    Dim i                        As Long           'Loop index
    Dim j                        As Long           'Loop index
    Dim nSubIdx                  As Long           'Subclass data index
    Dim sSubCode                 As String         'Subclass code string
    Const GMEM_FIXED             As Long = 0       'Fixed memory GlobalAlloc flag
    Const PAGE_EXECUTE_READWRITE As Long = &H40&
    Const PATCH_01               As Long = 18
    Const PATCH_02               As Long = 68      'Address of the previous WndProc
    Const PATCH_03               As Long = 78      'Relative address of SetWindowsLong
    Const PATCH_06               As Long = 116     'Address of the previous WndProc
    Const PATCH_07               As Long = 121     'Relative address of CallWindowProc
    Const PATCH_0A               As Long = 186     'Address of the owner object
    Const FUNC_CWP               As String = "CallWindowProcA"
    Const FUNC_EBM               As String = "EbMode"
    Const FUNC_SWL               As String = "SetWindowLongA"
    Const MOD_USER               As String = "user32"
    Const MOD_VBA5               As String = "vba5" 'Location of the EbMode function if running VB5
    Const MOD_VBA6               As String = "vba6" 'Location of the EbMode function if running VB6
    On Error Resume Next '
    m_Status = True
    m_LastHWND = lng_hWnd

    If sc_aBuf(1) = 0 Then
        sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E32F09C978078B450CF2AF75248D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF501CC3"
        i = 1

        Do While j < CODE_LEN
            j = j + 1
            sc_aBuf(j) = CByte("&H" & Mid$(sSubCode, i, 2))
            i = i + 2
        Loop

        If InIDE Then
            sc_aBuf(16) = &H90
            sc_aBuf(17) = &H90
            sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)

            If sc_pEbMode = 0 Then
                sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)
            End If
        End If

        zPatchVal VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me)
        sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
        sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
        ReDim sc_aSubData(0 To 0) As tSubData
    Else 'NOT SC_ABUF(1)...
        nSubIdx = zIdx(lng_hWnd, True)

        If nSubIdx = -1 Then
            nSubIdx = UBound(sc_aSubData()) + 1
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
        End If

        Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)
        VirtualProtect ByVal .nAddrSub, CODE_LEN, PAGE_EXECUTE_READWRITE, i
        RtlMoveMemory ByVal .nAddrSub, sc_aBuf(1), CODE_LEN
        .hwnd = lng_hWnd
        .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)
        zPatchRel .nAddrSub, PATCH_01, sc_pEbMode
        zPatchVal .nAddrSub, PATCH_02, .nAddrOrig
        zPatchRel .nAddrSub, PATCH_03, sc_pSWL
        zPatchVal .nAddrSub, PATCH_06, .nAddrOrig
        zPatchRel .nAddrSub, PATCH_07, sc_pCWP
    End With 'SC_ASUBDATA(NSUBIDX)

    On Error GoTo 0
End Function

Public Sub AddMsg(ByVal lng_hWnd As Long, ByVal uMSG As Long, Optional ByVal When As Integer = 2)

    With sc_aSubData(zIdx(lng_hWnd))

        If When And eMsgWhen.MSG_BEFORE Then
            zAddMsg uMSG, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub
        End If

        If When And eMsgWhen.MSG_AFTER Then
            zAddMsg uMSG, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub
        End If

    End With 'SC_ASUBDATA(ZIDX(LNG_HWND))

End Sub

Public Property Get AllowInIDE() As Boolean
    AllowInIDE = m_Allowinide
End Property

Public Property Let AllowInIDE(ByVal bValue As Boolean)
    m_Allowinide = bValue
End Property

Public Sub AttachAfterMSG(ByVal uMSG As Long)

    If m_LastHWND <> 0 Then
        AddMsg m_LastHWND, uMSG, 1
    End If

End Sub

Public Sub AttachMSG(ByVal uMSG As Long)

    If m_LastHWND <> 0 Then
        AddMsg m_LastHWND, uMSG, 2
    End If

End Sub

Private Sub Class_Initialize()
    m_Allowinide = False
End Sub

Public Sub DetachMSG()
    StopAll
End Sub

Private Function InIDE() As Boolean

    If Not m_Allowinide Then
        Debug.Assert zSetTrue(InIDE)
    End If

End Function

Private Sub StopAll()
    Dim i As Long
    On Error Resume Next

    If m_Status Then
        i = UBound(sc_aSubData())

        Do While i >= 0

            With sc_aSubData(i)

                If .hwnd <> 0 Then
                    StopWnd .hwnd
                End If

            End With 'SC_ASUBDATA(I)

            i = i - 1
        Loop

        m_Status = False
    End If

    On Error GoTo 0
End Sub

Private Sub StopWnd(ByVal lng_hWnd As Long)
    On Error Resume Next

    With sc_aSubData(zIdx(lng_hWnd))
        SetWindowLongA .hwnd, GWL_WNDPROC, .nAddrOrig
        zPatchVal .nAddrSub, PATCH_05, 0
        zPatchVal .nAddrSub, PATCH_09, 0
        GlobalFree .nAddrSub
        .hwnd = 0
        .nMsgCntB = 0
        .nMsgCntA = 0
        Erase .aMsgTblB
        Erase .aMsgTblA
    End With 'SC_ASUBDATA(ZIDX(LNG_HWND))

    On Error GoTo 0
End Sub

Private Sub zAddMsg(ByVal uMSG As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long   'Message table entry index
    Dim nOff1  As Long   'Machine code buffer offset 1
    Dim nOff2  As Long   'Machine code buffer offset 2
    On Error Resume Next '

    If uMSG = ALL_MESSAGES Then
        nMsgCnt = ALL_MESSAGES
    Else 'NOT UMSG...

        Do While nEntry < nMsgCnt
            nEntry = nEntry + 1

            If aMsgTbl(nEntry) = 0 Then
                aMsgTbl(nEntry) = uMSG
                Exit Sub
            ElseIf aMsgTbl(nEntry) = uMSG Then 'NOT AMSGTBL(NENTRY)...
                Exit Sub
            End If

        Loop

        nMsgCnt = nMsgCnt + 1
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
        aMsgTbl(nMsgCnt) = uMSG
    End If

    If When = eMsgWhen.MSG_BEFORE Then
        nOff1 = PATCH_04
        nOff2 = PATCH_05
    Else 'NOT WHEN...
        nOff1 = PATCH_08
        nOff2 = PATCH_09
    End If

    If uMSG <> ALL_MESSAGES Then
        zPatchVal nAddr, nOff1, VarPtr(aMsgTbl(1))
    End If

    zPatchVal nAddr, nOff2, nMsgCnt
    On Error GoTo 0
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    On Error Resume Next '
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc
End Function

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    On Error Resume Next '
    zIdx = UBound(sc_aSubData)

    Do While zIdx >= 0

        With sc_aSubData(zIdx)

            If .hwnd = lng_hWnd Then
                If Not bAdd Then
                    Exit Function
                End If

            ElseIf .hwnd = 0 Then 'NOT .HWND...

                If bAdd Then
                    Exit Function
                End If
            End If

        End With 'SC_ASUBDATA(ZIDX)

        zIdx = zIdx - 1
    Loop

    If Not bAdd Then
        Debug.Assert False
    End If

    On Error GoTo 0
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    RtlMoveMemory ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    RtlMoveMemory ByVal nAddr + nOffset, nValue, 4
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function

Private Sub Class_Terminate()
    StopAll
End Sub

Public Function IsInIDE() As Boolean
    Debug.Assert zSetTrue(IsInIDE)
End Function

Public Sub UnSubclass(ByVal lHwnd As Long)
    StopWnd lHwnd
End Sub

Public Function GetOriginalHwnd(ByVal lHwnd As Long) As Long
    GetOriginalHwnd = sc_aSubData(zIdx(lHwnd)).nAddrOrig
End Function
