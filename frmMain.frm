VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "Clipboard Logger"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Timer tmrMain 
      Interval        =   1000
      Left            =   3360
      Top             =   3120
   End
   Begin VB.ListBox lst 
      Height          =   3120
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "더블클릭해서 복사"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   3615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lst_DblClick()
On Error GoTo ErrHandler
    Clipboard.Clear
    Clipboard.SetText (lst.List(lst.ListIndex))
    MsgBox "복사됨", vbInformation, "Clipboard Logger"
Exit Sub
ErrHandler:
    MsgBox "에러", vbExclamation, "Clipboard Logger"
End Sub

Private Sub tmrMain_Timer()
On Error GoTo ErrHandler
Static clip As String
If clip <> Clipboard.GetText Then
    lst.AddItem (Clipboard.GetText)
    clip = Clipboard.GetText
End If
ErrHandler:
End Sub
