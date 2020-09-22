VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objclsResource As clsResource
Private Sub Command1_Click()
    'IMPORTANT THE TESTDLL.DLL RESOURCE HAS'NT BEEN PROVIDED JUST SHOWING YOU HOW TO USE
    'THE RESOURCE, Replace it by your own dll!
    Set objclsResource = New clsResource
    'load the dll
    If objclsResource.IntializeResource(App.Path & "\TESTRES.DLL") <> 0 Then
        'set mouse pointer to custom
        Me.MousePointer = 99
        'load the background bmp from resource
        Me.Picture = objclsResource.LoadGraphic(103, BITMAP_IMAGE)
        'load the mouse pointer from resource
        Set Me.MouseIcon = objclsResource.LoadGraphic(104, CURSOR_IMAGE)
        'play wave file from resource
        objclsResource.PlayWave 105
    End If
End Sub
