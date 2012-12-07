VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form INT900 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   "학생사진 업로드"
   ClientHeight    =   3000
   ClientLeft      =   7005
   ClientTop       =   4665
   ClientWidth     =   6570
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6555
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame2"
         Height          =   2925
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   6495
         Begin SHDocVwCtl.WebBrowser WB 
            Height          =   2505
            Left            =   60
            TabIndex        =   2
            Top             =   390
            Width           =   6390
            ExtentX         =   11271
            ExtentY         =   4419
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   ""
         End
         Begin VB.Label lblClose 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "닫 기"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   5160
            TabIndex        =   3
            Top             =   90
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "INT900"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'## 사진 업로드
'## 2007.12.12

Option Explicit

Private miInputValue    As Integer
Private msHtml          As String

Private sFileLocation   As String
Private sSchNO          As String

Private Const sHostName = "www.dshw.co.kr"
Private Const sPort = "80"

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
End Sub

Public Sub Save_Photo(ByVal aFileLocation As String, ByVal aSchNO As String)
    
    sFileLocation = aFileLocation
    sSchNO = aSchNO
    
    Call SetInputValue
    Call SetMsHtml
    
End Sub


Private Sub SetInputValue()
    Dim i1      As Integer
    Dim i10     As Integer
    Dim i100    As Integer
    Dim i1000   As Integer
    Dim iX      As Integer
    
    i1 = getRndNumber()
    i10 = getRndNumber()
    i100 = getRndNumber()
    i1000 = getRndNumber()

    i10 = i10 * 10
    i100 = i100 * 100
    i1000 = i1000 * 1000
    
    iX = i1 + i10 + i100 + i1000

    miInputValue = iX
End Sub

Function getRndNumber() As Integer
    Randomize
    Dim i       As Integer
    
    i = CInt(Int((9 * Rnd()) + 1))
    
    getRndNumber = i
End Function

Sub SetMsHtml()
    Dim sHTML       As String
    Dim sStdCD      As String
    Dim iOutValue   As Integer
    
    Dim sUrl        As String
    
    iOutValue = Int(miInputValue / 7)   '7로 나누고
    iOutValue = iOutValue * 3           '그에 3을 곱한다
    
    WB.Navigate "about:blank"
    
    If sPort = "80" Then
        sUrl = sHostName
    Else
        sUrl = sHostName & ":" & sPort
    End If
    '"http://" & sUrl & "/upload.php?"
    msHtml = "<form enctype=multipart/form-data action=" & _
            "http://" & sUrl & "/NDOC/CHECK/DSHW/upload.php?" & _
            "&sFile=" & sFileLocation & _
            "&sSCHNO=" & sSchNO & _
            "&val2=" & Trim(Str(miInputValue)) & _
            "&val3=" & Trim(Str(iOutValue)) & _
            " method=post>" & _
            "<center><input type=hidden name=MAX_FILE_SIZE value=500000>" & _
            "<input name=xfile type=file /><p>" & _
            "<font color=336699 size=2>400K 미만의 파일만 등록할 수 있습니다.<br>" & _
            "파일이름은 시스템에서 정의되어집니다.<p>" & _
            "<input type=submit value=전송하기 />" & _
            "</form>"
    
    DoEvents
    WB.Document.Write msHtml
    
End Sub

Private Sub lblClose_Click()
    Unload INT900
End Sub
