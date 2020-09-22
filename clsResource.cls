VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'declare our API's and CONSTANTS
Private Const dblID As Double = -2.51702880262616E-101
Private Const LR_DEFAULTCOLOR = &H0
Private Const IMAGE_BITMAP = 0
Private Const IMAGE_CURSOR = 2
Private Const IMAGE_ICON = 1
Private Const SND_APPLICATION = &H80         '  look for application specific association
Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_FILENAME = &H20000     '  name is a file name
Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Private Const SND_PURGE = &H40               '  purge non-static events for task
Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Private Const SND_SYNC = &H0
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PICDESC As PICDESC, RefIID As IID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
     Destination As Any, _
     Source As Any, _
     ByVal Length As Integer)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As Long, ByVal hModule As Long, ByVal dwFlags As Long) As Long


Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type
Private Type PICDESC
  cbSizeOfStruct As Long
  picType As Long
  hGdiObj As Long
  hPalOrXYExt As Long
End Type

Public Enum GraphicTypes
    BITMAP_IMAGE = 1
    ICON_IMAGE = 3
    CURSOR_IMAGE = 4
End Enum
Dim m_lngResourceHandle As Long
Public Function IntializeResource(strResourceName As String) As Integer
    'Load the dll in memory and get its handle
    m_lngResourceHandle = LoadLibrary(strResourceName)
    If m_lngResourceHandle = 0 Then IntializeResource = 0 Else IntializeResource = 1
End Function
Public Function LoadGraphic(lngGraphicID As Long, intGraphicType As GraphicTypes)
    Dim lngImageH As Long
    Dim stdp As StdPicture
    Select Case intGraphicType
        Case BITMAP_IMAGE
            'load the BMP Image and get BMP Handle
            lngImageH = LoadImage(m_lngResourceHandle, lngGraphicID, IMAGE_BITMAP, 0&, 0&, LR_DEFAULTCOLOR)
        Case ICON_IMAGE
            'load the Icon Image and get Icon Handle
            lngImageH = LoadImage(m_lngResourceHandle, lngGraphicID, IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR)
        Case CURSOR_IMAGE
            'load the Cursor Image and get Cursor Handle
            lngImageH = LoadImage(m_lngResourceHandle, lngGraphicID, IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR)
    End Select
    'Pass the handle to the create function to create a picture of type stdpicture
    CreateBitmapPicture lngImageH, intGraphicType, stdp
    'return the picture
    Set LoadGraphic = stdp
    'free our memory
    Set stdp = Nothing
End Function
Function CreateBitmapPicture(ByVal hBmp As Long, intGraphicType As GraphicTypes, stdp As StdPicture)
    Dim R As Long, Pic As PicBmp, IPic As StdPicture, IID_IDispatch As IID
    Dim lngPicType As Long, picdes As PICDESC
    
    Select Case intGraphicType
        'assign the respective GUID's and Values
        Case BITMAP_IMAGE
            lngPicType = vbPicTypeBitmap
            With IID_IDispatch
                .Data1 = &H20400
                .Data4(0) = &HC0
                .Data4(7) = &H46
            End With
            picdes.picType = vbPicTypeBitmap
        Case ICON_IMAGE
            lngPicType = vbPicTypeIcon
            With IID_IDispatch
                .Data1 = &H7BF80980
                .Data2 = &HBF32
                .Data3 = &H101A
            End With
            picdes.picType = vbPicTypeIcon
            MoveMemory IID_IDispatch.Data4(0), dblID, 8
        Case CURSOR_IMAGE
            lngPicType = vbPicTypeIcon
            With IID_IDispatch
                .Data1 = &H7BF80980
                .Data2 = &HBF32
                .Data3 = &H101A
            End With
            MoveMemory IID_IDispatch.Data4(0), dblID, 8
            picdes.picType = vbPicTypeIcon
    End Select
    
    picdes.cbSizeOfStruct = Len(picdes)
    picdes.hGdiObj = hBmp

    'Create the picture
    R = OleCreatePictureIndirect(picdes, IID_IDispatch, True, stdp)
End Function
Public Function PlayWave(lngWaveID As Long)
    'plays sound directly from resource
    'should be stored as WAVE not CUSTOM
    Debug.Print PlaySound(lngWaveID, m_lngResourceHandle, SND_RESOURCE Or SND_ASYNC)
End Function
Private Sub Class_Terminate()
    'free our loaded DLL
    FreeLibrary m_lngResourceHandle
End Sub
