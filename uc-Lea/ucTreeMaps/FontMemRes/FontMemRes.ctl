VERSION 5.00
Begin VB.UserControl FontMemRes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1095
   InvisibleAtRuntime=   -1  'True
   Picture         =   "FontMemRes.ctx":0000
   PropertyPages   =   "FontMemRes.ctx":1B42
   ScaleHeight     =   1095
   ScaleWidth      =   1095
   ToolboxBitmap   =   "FontMemRes.ctx":1B53
End
Attribute VB_Name = "FontMemRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------
'Autor: Leandro Ascierto
'Date: 30/12/2019
'Web: www.leandroascierto.com
'Module name: FontMemRes
'Gratitude: LaVolpe (funcions), Covein (funcions)
'Version: 1.0.0
'------------------------------------------------------
Private Const FR_PRIVATE           As Long = &H10
Private Const FR_NOT_ENUM          As Long = &H20

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function TlsGetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
Private Declare Function TlsSetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long, ByVal lpTlsValue As Long) As Long
Private Declare Function TlsFree Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
Private Declare Function TlsAlloc Lib "kernel32.dll" () As Long

Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipNewPrivateFontCollection Lib "GdiPlus.dll" (ByRef mFontCollection As Long) As Long
Private Declare Function GdipPrivateAddMemoryFont Lib "GdiPlus.dll" (ByVal mFontCollection As Long, ByRef mMemory As Any, ByVal mLength As Long) As Long
Private Declare Function GdipPrivateAddFontFile Lib "GdiPlus.dll" (ByVal mFontCollection As Long, ByVal mFilename As String) As Long
Private Declare Function GdipDeletePrivateFontCollection Lib "GdiPlus.dll" (ByRef mFontCollection As Long) As Long

Private Declare Function AddFontMemResourceEx Lib "gdi32.dll" (ByRef pvoid As Any, ByVal dword As Long, ByRef DESIGNVECTOR, ByRef pDword As Long) As Long
Private Declare Function RemoveFontMemResourceEx Lib "gdi32.dll" (ByVal fh As Long) As Long
Private Declare Function AddFontResourceExW Lib "gdi32.dll" (ByVal lpszFilename As Long, Optional ByVal fl As Long = FR_PRIVATE Or FR_NOT_ENUM, Optional ByVal pdv As Long) As Long
Private Declare Function RemoveFontResourceExW Lib "gdi32.dll" (ByVal lpFileName As Long, Optional ByVal fl As Long = FR_PRIVATE Or FR_NOT_ENUM, Optional ByVal pdv As Long) As Long

Private Const TLS_MINIMUM_AVAILABLE As Long = 64

Private Type GDIPlusStartupInput
    GdiPlusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
End Type

Private Type tFiles
    sName                       As String
    bvData()                    As Byte
End Type

Private m_UseGdiClassic As Boolean
Private m_UseGdiPlus As Boolean
Private hFontCollection As Long
Private GdipToken As Long
Private c_tvFiles()             As tFiles
Private c_bvData()              As Byte


Public Property Get UseGdiClassic() As Boolean
    UseGdiClassic = m_UseGdiClassic
End Property

Public Property Let UseGdiClassic(ByVal New_Value As Boolean)
    m_UseGdiClassic = New_Value
    PropertyChanged "UseGdiClassic"
End Property

Public Property Get UseGdiPlus() As Boolean
    UseGdiPlus = m_UseGdiPlus
End Property

Public Property Let UseGdiPlus(ByVal New_Value As Boolean)
    m_UseGdiPlus = New_Value
    PropertyChanged "UseGdiPlus"
End Property

Private Sub UserControl_InitProperties()
    m_UseGdiClassic = True
    m_UseGdiPlus = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        If CBool(.ReadProperty("bData", False)) Then
            c_bvData() = .ReadProperty("bvData")
            Call UnpackData(c_bvData)
        End If
    
        m_UseGdiClassic = .ReadProperty("UseGdiClassic", True)
        m_UseGdiPlus = .ReadProperty("UseGdiPlus", True)
    End With
    
    AddMemFonts
    
End Sub

Public Property Get Gdip_hFontCollection() As Long
    Gdip_hFontCollection = hFontCollection
End Property


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        If IsArrayDim(VarPtrArray(c_bvData)) Then
            Call .WriteProperty("bvData", c_bvData)
            Call .WriteProperty("bData", True)
        Else
            Call .WriteProperty("bData", False)
        End If
        Call .WriteProperty("UseGdiClassic", m_UseGdiClassic, True)
        Call .WriteProperty("UseGdiPlus", m_UseGdiPlus, True)
    End With
End Sub


Public Function AddMemFonts()
    Dim lngFontCount As Long, hFontMemRes As Long, i As Long
    Dim GdipStartupInput As GDIPlusStartupInput

    If Not IsArrayDim(VarPtrArray(c_tvFiles)) Then
        Exit Function
    End If
    'Don't worry, this function is only called once, both in the ide and compiled
    If m_UseGdiClassic Then
        If ReadValue(&HFCC) = 0 Then 'HFCC = Handle Font Colection Classic ;)
            For i = 0 To UBound(c_tvFiles)
                With c_tvFiles(i)
                    hFontMemRes = AddFontMemResourceEx(.bvData(0), UBound(.bvData) + 1, 0&, lngFontCount)
                End With
            Next
            WriteValue &HFCC&, hFontMemRes
        End If
    End If
    
    If m_UseGdiPlus Then
        If ReadValue(&HFC) <> 0 Then Exit Function 'HFC = Handle Font Colection :)
        GdipStartupInput.GdiPlusVersion = 1&
        If GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0) = 0 Then 'Gdip not Shutdown, will remain in memory until the process is finished
            If GdipNewPrivateFontCollection(hFontCollection) = 0 Then 'FontCollection is not downloaded, it remains in memory until the process is finished
                WriteValue &HFC&, hFontCollection
                For i = 0 To UBound(c_tvFiles)
                    With c_tvFiles(i)
                        GdipPrivateAddMemoryFont hFontCollection, .bvData(0), UBound(.bvData) + 1
                    End With
                Next
            End If
        End If
    End If

End Function


Friend Function ppgGetStream() As Byte()
    ppgGetStream = c_bvData
End Function

Friend Function ppgSetStream(ByRef bvData() As Byte)
    c_bvData = bvData
    Call PropertyChanged("bvData")
End Function

Private Sub UserControl_Resize()
    UserControl.Size 48 * Screen.TwipsPerPixelX, 48 * Screen.TwipsPerPixelY
End Sub

'==================================================================================
'////////////////////////////      HELPER FUNCTIONS      \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress    As Long
    
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Private Function UnpackData(ByRef bvData() As Byte) As Boolean
    Dim cBag        As New PropertyBag
    Dim i           As Long
    Dim lCount      As Long
    
    If Not IsArrayDim(VarPtrArray(bvData)) Then
        Exit Function
    End If
    
    With cBag
        .Contents = bvData
        lCount = .ReadProperty("Index", 0)
    
        If lCount = 0 Then Exit Function
        lCount = lCount - 1
    
        ReDim c_tvFiles(lCount)
    
        For i = 0 To lCount
            c_tvFiles(i).bvData = .ReadProperty("FILE_" & i)
            c_tvFiles(i).sName = .ReadProperty("NAME_" & i)
        Next
    End With
    
    UnpackData = True
End Function

'las dos funciones a continuacion son de cobein y con algunas modificaciones mias,
'las he utilizado para crear una bandera publica sin tener que agregar un modulo publico.
Private Function WriteValue(ByVal lProp As Long, ByVal lValue As Long) As Boolean
    Dim lFlagIndex As Long
    Dim i       As Long
    Dim lIndex  As Long: lIndex = -1
    
    For i = 0 To TLS_MINIMUM_AVAILABLE - 1
        If TlsGetValue(i) = lProp Then
            lIndex = i + 1
            Exit For
        End If
    Next

    If lIndex = -1 Then
        Do
            lFlagIndex = TlsAlloc '// Find two consecutive slots
            lIndex = TlsAlloc
            If lIndex >= TLS_MINIMUM_AVAILABLE Then Exit Function
        Loop While Not lFlagIndex + 1 = lIndex
        Call TlsSetValue(lFlagIndex, lProp)
        Call TlsSetValue(lIndex, lValue)
        WriteValue = True
    End If
End Function

'Autor: Cobein
Private Function ReadValue(ByVal lProp As Long, Optional Default As Long) As Long
    Dim i       As Long
    For i = 0 To TLS_MINIMUM_AVAILABLE - 1
        If TlsGetValue(i) = lProp Then
            ReadValue = TlsGetValue(i + 1)
            Exit Function
        End If
    Next
    ReadValue = Default
End Function
