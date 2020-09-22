VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBind 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exe Binder By: Odesa - www.odesayazilim.com"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
   Icon            =   "bind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "bind.frx":3CD2
   ScaleHeight     =   3870
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "bind.frx":45BC4
      Top             =   2760
      Width           =   6615
   End
   Begin EXE_Binder.Button cmdMake 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ButtonStyle     =   8
      ButtonTheme     =   7
      CaptionEffect   =   3
      BackColor       =   5737262
      BackColorPressed=   8441155
      BackColorHover  =   10485588
      BorderColor     =   5737262
      BorderColorPressed=   8441155
      BorderColorHover=   10485588
      Caption         =   "Bind File(s)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EXE_Binder.Button cmdRemove 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ButtonStyle     =   8
      ButtonTheme     =   6
      CaptionEffect   =   3
      BackColor       =   2504331
      BackColorPressed=   4349166
      BackColorHover  =   4678655
      BorderColor     =   2504331
      BorderColorPressed=   4349166
      BorderColorHover=   4678655
      Caption         =   "Remove Exe File(s)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EXE_Binder.Button cmdAdd 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ButtonStyle     =   8
      ButtonStyleColors=   3
      ButtonTheme     =   5
      CaptionEffect   =   3
      BackColor       =   181749
      BackColorPressed=   11530238
      BackColorHover  =   14875135
      BorderColor     =   181749
      BorderColorPressed=   11530238
      BorderColorHover=   14875135
      Caption         =   "Add Exe File(S)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1395
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog File1 
      Left            =   1560
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Exe files (*.exe)|*.exe"
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1140
      Left            =   4440
      Picture         =   "bind.frx":45CE4
      Top             =   360
      Width           =   2310
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Odesa eXe Binder - Source Code:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label tag1 
      BackColor       =   &H80000012&
      Caption         =   "Binded EXE(s) Files list:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmBind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' eXe Binder Coded By: Alper ESKIKILIC - www.odesayazilim.com - odysseydhtbluefire@hotmail.com
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private fPaths() As String
Private fIndex As Integer
Private fPoint As Long

Private Sub Form_Load()
ReDim fPaths(10)
End Sub
Private Sub cmdAdd_Click()
File1.Flags = 4096 Or 4: File1.ShowOpen
If File1.FileName = vbNullString Then Exit Sub

lstFiles.AddItem File1.FileTitle
fPaths(fIndex) = File1.FileName
fIndex = fIndex + 1: File1.FileName = ""

If fIndex > UBound(fPaths) Then
ReDim Preserve fPaths(fIndex + 10)
End If
End Sub

Private Sub cmdRemove_Click()
Dim X As Long, OldIndex As Long
If lstFiles.ListIndex = -1 Then Exit Sub
OldIndex = lstFiles.ListIndex
lstFiles.RemoveItem (OldIndex)

For X = OldIndex To lstFiles.ListCount
    fPaths(X) = fPaths(X + 1)
Next X: fIndex = fIndex - 1
End Sub

Private Sub cmdMake_Click()
Dim Buffer As String, X As Long, Path As String: Path = App.Path
On Error Resume Next: Kill (Path & "\output.exe"): DoEvents

If lstFiles.ListCount = 0 Then Exit Sub
Open Path & "\header.dat" For Binary As #1
    fPoint = LOF(1): Buffer = Space$(fPoint)
    Get #1, , Buffer
Close #1: fPoint = fPoint + 1
Open Path & "\output.exe" For Binary As #1
Put #1, , Buffer

For X = 0 To lstFiles.ListCount - 1
    Put #1, fPoint, strFile(X)
Next X: Close #1: Me.Caption = "Made@ " & Time

End Sub

Private Function strFile(Index As Long) As String
Dim hexSize As String, lof2 As Long

Open fPaths(Index) For Binary As #2
lof2 = LOF(2): hexSize = Hex(lof2)

    Do Until Len(hexSize) = 6
    hexSize = "0" & hexSize
    Loop
                                                

strFile = Space$(lof2)
fPoint = fPoint + lof2 + 6
Get #2, , strFile: Close #2

strFile = hexSize & strFile
End Function
