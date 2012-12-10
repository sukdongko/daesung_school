VERSION 5.00
Begin VB.Form LOG001 
   BorderStyle     =   1  '���� ����
   Caption         =   "�뼺�п� ���л���. �ݹ���. �ð�ǥ ���α׷�"
   ClientHeight    =   2520
   ClientLeft      =   7620
   ClientTop       =   6375
   ClientWidth     =   3990
   Icon            =   "LOG001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3990
   Begin VB.CommandButton cmdExit 
      Caption         =   "������"
      Height          =   435
      Left            =   2010
      TabIndex        =   3
      Top             =   1560
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "����"
      Height          =   435
      Left            =   690
      TabIndex        =   2
      Top             =   1560
      Width           =   1065
   End
   Begin VB.CommandButton cmdSchool 
      Caption         =   "�п� �����ϱ�"
      Height          =   300
      Left            =   3360
      TabIndex        =   8
      Top             =   270
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ComboBox cboSchool 
      Height          =   300
      Left            =   1260
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   4
      Top             =   180
      Width           =   1875
   End
   Begin VB.TextBox txtPass 
      BorderStyle     =   0  '����
      Height          =   300
      IMEMode         =   3  '��� ����
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "txtPass"
      Top             =   1035
      Width           =   1485
   End
   Begin VB.TextBox txtNM 
      BorderStyle     =   0  '����
      Height          =   300
      IMEMode         =   10  '�ѱ� 
      Left            =   1290
      MaxLength       =   50
      TabIndex        =   0
      Text            =   "txtNM"
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label lblShow 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "."
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3840
      TabIndex        =   9
      Top             =   2340
      Width           =   135
   End
   Begin VB.Label lblSchool 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "�п�����"
      Height          =   210
      Left            =   240
      TabIndex        =   7
      Top             =   270
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "��й�ȣ"
      Height          =   210
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "���"
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   645
      Width           =   975
   End
End
Attribute VB_Name = "LOG001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : LOG001
'   �� ��  �� �� : LOGIN ó��
'
'   ��   ��   �� : 2007/08/20
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################


Option Explicit
Private sini_Path      As String    '>> �뼺�п�
Private Const sDebug = 1


Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    Dim sData               As String * 255
    Dim sGbn                As String
    Dim nRtn                As Long
    
    Dim sTmp                As String
    Dim sDB_Tns_Location    As String
    Dim bFirstLogin         As Boolean
    
    bFirstLogin = False
    
    '## ���α׷� �������� ��� �� �ֵ��� ��. ��, update�� �ȵ�.
    If App.PrevInstance = False Then
        Call Update
    End If
    
    Me.KeyPreview = True
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Tag = "LOAD"
    
    
    txtNM.Text = ""
    txtPass.Text = ""
    
    If sDebug = 1 Then
        txtNM.Text = "ADMIN"
        txtPass.Text = "1"
    End If
    
    If sDebug = 0 Then
        cboSchool.Visible = False
        lblSchool.Visible = False
    End If
    
    basFunction.RemoveContextMenu txtNM
    basFunction.RemoveContextMenu txtPass
    
    With cboSchool
        .Clear
        .AddItem "�뷮��" & Space(30) & "N"
        .AddItem "����" & Space(30) & "K"
        .AddItem "����" & Space(30) & "S"
        .AddItem "���� M" & Space(30) & "P"
        
        .AddItem "���� M" & Space(30) & "M"
        .AddItem "�ָ����Ǵ�" & Space(30) & "W"
        .AddItem "�߰����Ǵ�" & Space(30) & "Q"
        
        .AddItem "����" & Space(30) & "J"
        .AddItem "�λ�" & Space(30) & "B"
        
        .ListIndex = 0
    End With
    
    ' ���� ���� ���� INI���� ��������.
    
    '>> ���α׷� INI ����
    sini_Path = App.Path & "\DAESUNG.INI"
    If Dir(sini_Path) = "" Then                                     '<< ������ ������ ����
        Call Create_School_ini_File
        '�����̾����� ù�α������� ó���Ѵ�. �п�����
        bFirstLogin = True
    End If
    
    '>>>>>>>>>>>>>>>>>>>>>> ���α׷� �������� ����
    sGbn = "SCHOOL"
    sData = ""
    nRtn = basModule.GetPrivateProfileString(sGbn, "SCHOOL", "", sData, 255, sini_Path)         '>> �б��ڵ�
    basModule.schcd = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
    If nRtn = 0 Then
        Call Create_School_ini_File
    End If
    
    
    Select Case Trim(basModule.schcd)
        Case "N"
            cboSchool.ListIndex = 0
        Case "K"
            cboSchool.ListIndex = 1
        Case "S"
            cboSchool.ListIndex = 2
        Case "P"
            cboSchool.ListIndex = 3
        Case "M"
            cboSchool.ListIndex = 4
            
        Case "W"
            cboSchool.ListIndex = 5
        Case "Q"
            cboSchool.ListIndex = 6
            
        Case "J"
            cboSchool.ListIndex = 7
        Case "B"
            cboSchool.ListIndex = 8
        
        Case Else
            cboSchool.ListIndex = 0
    End Select
    
    
    sData = ""
    nRtn = basModule.GetPrivateProfileString(sGbn, "SCHOOL_NM", "", sData, 255, sini_Path)      '>> ������ ��� �ڵ�
    basModule.SchNM = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
    
    
    If nRtn = 0 Then
        Call Create_School_ini_File
    End If
    
    sData = ""
    nRtn = basModule.GetPrivateProfileString(sGbn, "DB", "", sData, 255, sini_Path)             '>> DB����
    basModule.connDB = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
    If nRtn = 0 Then
        Call Create_School_ini_File
    End If
    
    
    '�� �п��� ���� �����ڵ� ���� ����
    Call basGwamok.setConstant
        
    
    
    '## ���ӵ����� ó��
    '## DB ���� : ������ �Ϸ�Ǹ� => DBConn �� connection ������ �־����ϴ�.
    If basDataBase.DataBase_Connection() = False Then
        MsgBox "���� ���� �����ڿ��� ���� �ٶ��ϴ�"
        Exit Sub
    End If
            
    Me.Tag = ""
    
    
    
    If True = bFirstLogin Then
        cboSchool.Visible = True
        lblSchool.Visible = True
    End If
    
End Sub

Private Sub Create_School_ini_File()
    Dim sGbn        As String
    Dim nRtn        As Long
    
    basModule.schcd = "N"
    basModule.SchNM = "�뷮��"
    basModule.connDB = "MIMAC"
        
    sGbn = "SCHOOL"
    'nRtn = basModule.WritePrivateProfileString(sGbn, "PATH_ORACLE_TNS", basDataBase.TNS_Path1, sini_Path)                  '<< oracle tns ��� - ������ ����ɼ��ִ�.
    nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", schcd, sini_Path)                  '<< �п�
    nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", SchNM, sini_Path)          '<< �п���
    nRtn = basModule.WritePrivateProfileString(sGbn, "DB", connDB, sini_Path)                  '<< DB ���� - mimac �Ǽ���, dev ���߼���
        
        
End Sub


Private Sub cboSchool_Click()
    If StrComp(Trim(Me.Tag), "LOAD", vbTextCompare) = 0 Then Exit Sub
    
    Call cmdSchool_Click
End Sub

'>> �п�����
Private Sub cmdSchool_Click()
    Dim sGbn        As String
    Dim nRtn        As Long
    
    If StrComp(Trim(Me.Tag), "LOAD", vbTextCompare) = 0 Then Exit Sub
    
    Select Case Trim(Right(cboSchool.Text, 30))
        Case "N"
            If MsgBox("�뷮�� �п��Դϴ�." & vbCrLf & "�½��ϱ�?", vbQuestion + vbYesNo, "�п�����") = vbNo Then Exit Sub
        Case "K"
            If MsgBox("���� �п��Դϴ�." & vbCrLf & "�½��ϱ�?", vbQuestion + vbYesNo, "�п�����") = vbNo Then Exit Sub
        Case "S"
            If MsgBox("���� �п��Դϴ�." & vbCrLf & "�½��ϱ�?", vbQuestion + vbYesNo, "�п�����") = vbNo Then Exit Sub
        Case "P"
            If MsgBox("���ĸ��̸� �п��Դϴ�." & vbCrLf & "�½��ϱ�?", vbQuestion + vbYesNo, "�п�����") = vbNo Then Exit Sub
        Case "M"
            If MsgBox("�������̸� �п��Դϴ�." & vbCrLf & "�½��ϱ�?", vbQuestion + vbYesNo, "�п�����") = vbNo Then Exit Sub
            
        Case "W"
            If MsgBox("�ָ����Ǵ� �п��Դϴ�." & vbCrLf & "�½��ϱ�?", vbQuestion + vbYesNo, "�п�����") = vbNo Then Exit Sub
        Case "Q"
            If MsgBox("�߰����Ǵ� �п��Դϴ�." & vbCrLf & "�½��ϱ�?", vbQuestion + vbYesNo, "�п�����") = vbNo Then Exit Sub
            
        Case "J"
            If MsgBox("���� �п��Դϴ�." & vbCrLf & "�½��ϱ�?", vbQuestion + vbYesNo, "�п�����") = vbNo Then Exit Sub
        Case "B"
            If MsgBox("�λ� �п��Դϴ�." & vbCrLf & "�½��ϱ�?", vbQuestion + vbYesNo, "�п�����") = vbNo Then Exit Sub
            
    End Select
    
    sGbn = "SCHOOL"
    Select Case Trim(Right(cboSchool.Text, 30))
        Case "N"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "N", sini_Path)                  '<< �п�
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "�뷮��", sini_Path)          '<< �п���
            
            schcd = "N":    SchNM = "�뷮��"
        Case "K"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "K", sini_Path)                  '<< �п�
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "����", sini_Path)            '<< �п���
            
            schcd = "K":    SchNM = "����"
        Case "S"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "S", sini_Path)                  '<< �п�
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "����", sini_Path)            '<< �п���
            
            schcd = "S":    SchNM = "����"
        Case "P"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "P", sini_Path)                  '<< �п�
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "���� M", sini_Path)          '<< �п���
            
            schcd = "P":    SchNM = "���� M"
        Case "M"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "M", sini_Path)                  '<< �п�
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "���� M", sini_Path)          '<< �п���
            
            schcd = "M":    SchNM = "���� M"
            
        Case "W"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "W", sini_Path)                  '<< �п�
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "�ָ����Ǵ�", sini_Path)      '<< �п���
            
            schcd = "W":    SchNM = "�ָ����Ǵ�"
        Case "Q"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "Q", sini_Path)                  '<< �п�
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "�߰����Ǵ�", sini_Path)      '<< �п���
            
            schcd = "Q":    SchNM = "�߰����Ǵ�"
            
        Case "J"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "J", sini_Path)                  '<< �п�
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "����", sini_Path)        '<< �п���
            
            schcd = "J":    SchNM = "����"
            
        Case "B"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "B", sini_Path)                  '<< �п�
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "�λ�", sini_Path)            '<< �п���
            
            schcd = "B":    SchNM = "�λ�"
            
    End Select
    
    MsgBox "�Ϸ��Ͽ����ϴ�.", vbInformation + vbOKOnly, "�п�����"
    
End Sub



'>> ���α׷� ��뿩��
Private Sub cmdOK_Click()
    Dim sSql        As String
    Dim sTmp        As String
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim ni          As Long
    
    Dim bChk        As Boolean
    
    If Trim(txtNM.Text) = "" Then
        MsgBox "������� ��������.", vbExclamation + vbOKOnly, "Ȯ��"
        Exit Sub
    End If
    
    If Trim(txtPass.Text) = "" Then
        MsgBox "��й�ȣ�� ��������.", vbExclamation + vbOKOnly, "Ȯ��"
        Exit Sub
    End If
    
    bChk = False
    
    '>> ȸ��Ȯ��
    If Trim(txtNM.Text) = "ADMIN" Then
        sSql = ""
        sSql = sSql & " SELECT EMPNO, EMPNM, PASSWD "
        sSql = sSql & "   FROM CLEMP01TB "
        sSql = sSql & "  WHERE EMPNM  = '" & Trim(txtNM.Text) & "' "
        sSql = sSql & "    AND PASSWD = '" & Trim(txtPass.Text) & "' "
    Else
        sSql = ""
        sSql = sSql & " SELECT EMPNO, EMPNM, PASSWD "
        sSql = sSql & "   FROM CLEMP01TB "
        sSql = sSql & "  WHERE ACID   = '" & Trim(basModule.schcd) & "' "
        sSql = sSql & "    AND EMPNM  = '" & Trim(txtNM.Text) & "' "
        sSql = sSql & "    AND PASSWD = '" & Trim(txtPass.Text) & "' "
    End If
    
    On Error GoTo ErrStmt
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    
    With DBCmd
        .ActiveConnection = basDataBase.DBConn
        .CommandTimeout = 60
        .CommandText = sSql
        .CommandType = adCmdText
        
    End With
    
    With DBRec
        .Open DBCmd, , adOpenStatic, adLockReadOnly, -1         '<< dynamic ���·� �����Ǹ� record count�� �� �� ����.
        Do While .State And adStateExecuting
            DoEvents
        Loop
        
        If .RecordCount = 1 Then
            If StrComp(Trim(txtNM.Text), .Fields("EMPNM"), vbTextCompare) = 0 And _
               StrComp(Trim(txtPass.Text), .Fields("PASSWD"), vbTextCompare) = 0 Then
                        
                bChk = True         '<< Ȯ�� OK
                
                basModule.RegID = .Fields("EMPNO")
                
            End If
        Else
            MsgBox "����ڰ� �����ϴ�.", vbExclamation + vbOKOnly, "LOGIN"
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    If bChk = True Then
        Load MDI001
        MDI001.WindowState = 2
        MDI001.Show
        
        Unload LOG001
    End If
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    'MsgBox "���α׷� ���üũ�� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description & vbCrLf & _
           basDataBase.DBConn, vbCritical + vbOKOnly, "LOGIN"
    
    MsgBox "���α׷� ���üũ�� ������ �߻��Ͽ����ϴ�.  " & vbCrLf & _
            "ȯ�溯�� Path�� ����Ŭ ��θ� Ȯ���ϼ���." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description & vbCrLf _
           , vbCritical + vbOKOnly, "LOGIN"
           
    On Error GoTo 0
End Sub



'>> liveupdate ó��
Private Sub Update()
    If Command = "no_update" Then
        Call basModule.SleepEx(1000, 0)
        Call RenameUpdateExe
    Else
        Call RunUpdateExe
    End If
End Sub

Private Sub RunUpdateExe()
    On Error GoTo EH
    Call Shell(App.Path & "\update.exe " & App.EXEName & ",�뼺�п� �ݹ��� ���α׷���", vbNormalFocus)
    End
EH:
    MsgBox "���̺������Ʈ�� �����Ǿ����ϴ�.", vbExclamation, "�뼺�п� �ݹ��� ���α׷�"
End Sub

Sub RenameUpdateExe()
    On Error Resume Next

    
    Call Shell(Environ("COMSPEC") & " /c " & _
            "move /y " & Chr(34) & App.Path & "\update" & "\update_x.exe" & Chr(34) & " " & Chr(34) & App.Path & "\update.exe" & Chr(34), vbNormalFocus)

    Call Kill(App.Path & "\update_x.exe")
End Sub



Private Sub lblShow_Click()
    cboSchool.Visible = True
    lblSchool.Visible = True
End Sub

Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call cmdOK_Click
            
    End Select
End Sub
