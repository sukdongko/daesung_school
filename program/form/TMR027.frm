VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form TMR027 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  '���� ����
   Caption         =   "�ð�ǥ ����� >> �̵����� �ð�ǥ ������"
   ClientHeight    =   4695
   ClientLeft      =   1680
   ClientTop       =   6270
   ClientWidth     =   13455
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   13455
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  '����
      Caption         =   "Frame2"
      Height          =   4425
      Left            =   30
      TabIndex        =   17
      Top             =   30
      Width           =   13395
      Begin VB.ComboBox cboKaeyol 
         Height          =   300
         Left            =   180
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   1
         Top             =   90
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�ݺ� ���񳻿� ����ϱ�"
         Height          =   435
         Left            =   10680
         TabIndex        =   16
         Top             =   3840
         Width           =   2475
      End
      Begin VB.ComboBox cboLsnType 
         Height          =   300
         Left            =   1500
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   2
         Top             =   90
         Width           =   975
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "�ݺ� ���񳻿� ��ȸ�ϱ�"
         Height          =   435
         Left            =   2940
         TabIndex        =   0
         Top             =   60
         Width           =   2475
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Caption         =   "Frame1"
         Height          =   435
         Left            =   30
         TabIndex        =   18
         Top             =   540
         Width           =   13335
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            Height          =   240
            Index           =   11
            Left            =   12270
            TabIndex        =   14
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H000000FF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   10
            Left            =   11160
            TabIndex        =   13
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H0000C000&
            Caption         =   "Option1"
            Height          =   240
            Index           =   9
            Left            =   10050
            TabIndex        =   12
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00FFFF00&
            Caption         =   "Option1"
            Height          =   240
            Index           =   8
            Left            =   8940
            TabIndex        =   11
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00FF80FF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   7
            Left            =   7830
            TabIndex        =   10
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Option1"
            Height          =   240
            Index           =   6
            Left            =   6720
            TabIndex        =   9
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H0000C0C0&
            Caption         =   "Option1"
            Height          =   240
            Index           =   5
            Left            =   5610
            TabIndex        =   8
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Option1"
            Height          =   240
            Index           =   4
            Left            =   4500
            TabIndex        =   7
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   3
            Left            =   3390
            TabIndex        =   6
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   2
            Left            =   2280
            TabIndex        =   5
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   1
            Left            =   1170
            TabIndex        =   4
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00FFFFFF&
            Caption         =   "���񳻿�"
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   90
            Width           =   1125
         End
      End
      Begin FPSpread.vaSpread sprGwamok 
         Height          =   2715
         Left            =   30
         TabIndex        =   15
         Top             =   990
         Width           =   13335
         _Version        =   393216
         _ExtentX        =   23521
         _ExtentY        =   4789
         _StockProps     =   64
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxRows         =   4
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "TMR027.frx":0000
      End
   End
End
Attribute VB_Name = "TMR027"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : TRM027
'   �� ��  �� �� :
'
'   ��   ��   �� : 2008/01/07
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit

Private Sub Form_Activate()
    sprGwamok.SetFocus
    If sprGwamok.MaxCols > 1 Then sprGwamok.SetActiveCell 1, 1
    
End Sub

Private Sub Form_Load()
    
    Me.Move 200, 900, 13530, 4790
    
    With sprGwamok
        .ShadowColor = basModule.ShadowColor2
        .ShadowDark = basModule.ShadowDark2
        .ShadowText = basModule.ShadowText2
        .GridColor = basModule.GridColor2
        .GrayAreaBackColor = basModule.GrayAreaBackColor2
        
        .MaxCols = 0
        .Row = SpreadHeader:        .RowHidden = True
        .Row = SpreadHeader + 1:    .RowHeight(.Row) = 16
        .Row = SpreadHeader + 2:    .RowHeight(.Row) = 16
        
    End With
        
    With cboKaeyol
        .Clear
        .AddItem "�ι�" & Space(30) & "01"
        .AddItem "�ڿ�" & Space(30) & "02"
        
        .ListIndex = 0
    End With
    
    With cboLsnType
        .Clear
        .AddItem "A type" & Space(30) & "A"
        .AddItem "B type" & Space(30) & "B"
        .AddItem "C type" & Space(30) & "C"
        
        .ListIndex = 0
    End With
    
    Call Find_LsnCD         '< �� ��ȸ
    
        
        
    cmdFind.Tag = ""
    
End Sub

Public Sub init_Data(ByVal aKaeyol As String, ByVal aLsnType As String)
    
    Me.Tag = "LOAD"
    
    sprGwamok.MaxCols = 0
    
    With cboKaeyol
        Select Case aKaeyol
            Case "01"
                .ListIndex = 0
                
                optTamgu(0).Caption = "����/����"
                optTamgu(1).Caption = "����":           optTamgu(1).Tag = "01"
                optTamgu(2).Caption = "����":           optTamgu(2).Tag = "02"
                optTamgu(3).Caption = "����":           optTamgu(3).Tag = "03"
                optTamgu(4).Caption = "�ѱ�":           optTamgu(4).Tag = "04"
                optTamgu(5).Caption = "�����":         optTamgu(5).Tag = "05"
                optTamgu(6).Caption = "����":           optTamgu(6).Tag = "06"
                optTamgu(7).Caption = "����":           optTamgu(7).Tag = "07"
                optTamgu(8).Caption = "��ġ":           optTamgu(8).Tag = "08"
                optTamgu(9).Caption = "�繮":           optTamgu(9).Tag = "09":             optTamgu(9).Visible = True
                optTamgu(10).Caption = "����":          optTamgu(10).Tag = "10":            optTamgu(10).Visible = True
                optTamgu(11).Caption = "����":          optTamgu(11).Tag = "11":            optTamgu(11).Visible = True
                
                optTamgu(0).Value = True            '�⺻����
                
            Case "02"
                .ListIndex = 1
                
                optTamgu(0).Caption = "����/����"
                optTamgu(1).Caption = "��1":            optTamgu(1).Tag = "51"
                optTamgu(2).Caption = "ȭ1":            optTamgu(2).Tag = "52"
                optTamgu(3).Caption = "��1":            optTamgu(3).Tag = "53"
                optTamgu(4).Caption = "��1":            optTamgu(4).Tag = "54"
                optTamgu(5).Caption = "��2":            optTamgu(5).Tag = "55"
                optTamgu(6).Caption = "ȭ2":            optTamgu(6).Tag = "56"
                optTamgu(7).Caption = "��2":            optTamgu(7).Tag = "57"
                optTamgu(8).Caption = "��2":            optTamgu(8).Tag = "58"
                
                optTamgu(9).Caption = "":               optTamgu(9).Tag = "00":             optTamgu(9).Visible = False
                optTamgu(10).Caption = "":              optTamgu(10).Tag = "00":            optTamgu(10).Visible = False
                optTamgu(11).Caption = "":              optTamgu(11).Tag = "00":            optTamgu(11).Visible = False
    
                optTamgu(0).Value = True            '�⺻����
                
        End Select
    End With
    
    With cboLsnType
        Select Case aLsnType
            Case "A"
                .ListIndex = 0
            Case "B"
                .ListIndex = 1
            Case "C"
                .ListIndex = 2
        End Select
    End With

    
    cmdFind.Tag = "FIRST"
        Call cmdFind_Click
        
    cmdFind.Tag = ""

    Me.Tag = ""

End Sub

Private Sub Find_LsnCD()
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nColor      As Long
    
    sprGwamok.MaxCols = 0
    
    On Error GoTo ErrStmt
    
    sStr = ""
'    sStr = sStr & "  SELECT LSNCD, LSNNM, LSNCDNM, LSN_CL"
'    sStr = sStr & "    FROM (SELECT *"
'    sStr = sStr & "            From SDLSN01TB"
'    sStr = sStr & "           WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
'    sStr = sStr & "             AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
'    sStr = sStr & "           ORDER BY LSNCDNM"
'    sStr = sStr & "          )"
'    sStr = sStr & "  Union All"
'    sStr = sStr & "  SELECT LSNCD, LSNNM, LSNCDNM, LSN_CL"
'    sStr = sStr & "    FROM (SELECT *"
'    sStr = sStr & "            From SDLSN02TB"
'    sStr = sStr & "           WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
'    sStr = sStr & "             AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
'    sStr = sStr & "           ORDER BY LSNCDNM"
'    sStr = sStr & "          )"

    sStr = sStr & "    SELECT A.LSNCD, LSNNM, LSNCDNM, LSN_CL "
    sStr = sStr & "      FROM (SELECT ACID, LSNCD, LSNNM, LSNCDNM, LSN_CL"
    sStr = sStr & "              FROM (SELECT *"
    sStr = sStr & "                      From SDLSN01TB"
    sStr = sStr & "                     WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                       AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "                     ORDER BY LSNCDNM"
    sStr = sStr & "                    )"
    sStr = sStr & "            Union All"
    sStr = sStr & "            SELECT ACID, LSNCD, LSNNM, LSNCDNM, LSN_CL"
    sStr = sStr & "              FROM (SELECT *"
    sStr = sStr & "                      From SDLSN02TB"
    sStr = sStr & "                     WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                       AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "                     ORDER BY LSNCDNM"
    sStr = sStr & "                    )"
    sStr = sStr & "            ) A, "
    sStr = sStr & "            SDLSN05TB B"
    sStr = sStr & "      WHERE A.ACID    = B.ACID"
    sStr = sStr & "        AND A.LSNCD   = B.LSNCD "
    sStr = sStr & "        AND A.ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "        AND B.LSNTYPE = '" & Trim(Right(cboLsnType.Text, 30)) & "'"
    
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    


    
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
       
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprGwamok.MaxCols = sprGwamok.MaxCols + 1
                sprGwamok.Col = sprGwamok.MaxCols
                
                sprGwamok.Row = SpreadHeader
                    sTmp = "":      If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD")):         sprGwamok.Text = sTmp
                sprGwamok.Row = SpreadHeader + 1
                    sTmp = "":      If IsNull(.Fields("LSNCDNM")) = False Then sTmp = Trim(.Fields("LSNCDNM")):     sprGwamok.Text = sTmp
                sprGwamok.Row = SpreadHeader + 2
                    sTmp = "":      If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM")):         sprGwamok.Text = sTmp
                    
                    
                sprGwamok.Row = 1:      Call basFunction.Set_SprType_Text(sprGwamok, "center", "center", 10, "")
                sprGwamok.Row = 2:      Call basFunction.Set_SprType_Text(sprGwamok, "center", "center", 10, "")
                sprGwamok.Row = 3:      Call basFunction.Set_SprType_Text(sprGwamok, "center", "center", 10, "")
                sprGwamok.Row = 4:      Call basFunction.Set_SprType_Text(sprGwamok, "center", "center", 10, "")
                
                .MoveNext       '<< �����׸�
                
            Next nRec
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�� ��ȸ"
End Sub

Private Sub cmdFind_Click()
    cmdFind.Enabled = False
        
        Call Find_Gwamok_Detail
        
    cmdFind.Enabled = True
End Sub


Private Sub Find_Gwamok_Detail()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim nCol        As Long
    Dim sLsnCD      As String
    Dim nColor      As Long
    
    Call Find_LsnCD         '< �� ��ȸ
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT LSNCD, ORD,"

    sStr = sStr & "         CASE WHEN      TRIM(SUBJCD) = '01' THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '02' THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '03' THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '04' THEN '�ѱ�'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '05' THEN '�����'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '06' THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '07' THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '08' THEN '��ġ'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '09' THEN '�繮'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '10' THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '11' THEN '����'"
    
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '51' THEN '��1'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '52' THEN 'ȭ1'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '53' THEN '��1'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '54' THEN '��1'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '55' THEN '��2'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '56' THEN 'ȭ2'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '57' THEN '��2'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '58' THEN '��2'"
    
    sStr = sStr & "         END END END END END END END END END END END"
    sStr = sStr & "         END END END END END END END END AS SUBJCD,"
    
    sStr = sStr & "         CASE WHEN      TRIM(SUBJCD) = '01' THEN 1"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '02' THEN 2"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '03' THEN 3"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '04' THEN 4"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '05' THEN 5"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '06' THEN 6"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '07' THEN 7"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '08' THEN 8"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '09' THEN 9"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '10' THEN 10"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '11' THEN 11"
    
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '51' THEN 1"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '52' THEN 2"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '53' THEN 3"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '54' THEN 4"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '55' THEN 5"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '56' THEN 6"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '57' THEN 7"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '58' THEN 8"
    sStr = sStr & "         ELSE 0"
    sStr = sStr & "         END END END END END END END END END END END"
    sStr = sStr & "         END END END END END END END END AS COLORS"
    
    sStr = sStr & "    FROM SDLSN06TB"
    sStr = sStr & "   WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND KAEYOL  = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "     AND LSNTYPE = '" & Trim(Right(cboLsnType.Text, 30)) & "'"
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    


'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
       
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                
                
                sLsnCD = "":      If IsNull(.Fields("LSNCD")) = False Then sLsnCD = Trim(.Fields("LSNCD"))
                If sLsnCD <> "" Then
                        
                    sprGwamok.Row = SpreadHeader
                    For nCol = 1 To sprGwamok.MaxCols Step 1
                        sprGwamok.Col = nCol
                        
                        If StrComp(Trim(sprGwamok.Text), sLsnCD, vbTextCompare) = 0 Then            '< LSNCD ��
                            If IsNumeric(.Fields("ORD")) = True Then                                '< ORD : ��
                                sprGwamok.Row = CLng(.Fields("ORD"))
                                
                                sTmp = "":      If IsNull(.Fields("SUBJCD")) = False Then sTmp = Trim(.Fields("SUBJCD"))
                                If sTmp = "" Then
                                    Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, "")
                                Else
                                    Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, sTmp)
                                End If
                                
                                sprGwamok.Row2 = sprGwamok.Row
                                sprGwamok.Col2 = sprGwamok.Col
                                sprGwamok.BlockMode = True
                                    nTmp = &HFFFFFF:        If IsNumeric(.Fields("COLORS")) = True Then nTmp = CLng(.Fields("COLORS"))
                                    If nTmp = 0 Or nTmp = &HFFFFFF Then
                                        nTmp = 0
                                        sprGwamok.BackColor = optTamgu(nTmp).BackColor
                                    Else
                                        sprGwamok.BackColor = optTamgu(nTmp).BackColor
                                    End If
                                    sprGwamok.BackColorStyle = BackColorStyleUnderGrid
                                sprGwamok.BlockMode = False
                                
                            End If
                        End If
                    Next nCol
                End If
                
                .MoveNext       '<< �����׸�
                
            Next nRec
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    If cmdFind.Tag = "" Then
        MsgBox "��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "���� ��ϳ��� ��ȸ"
    End If
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "���� ��ϳ��� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "���� ��ϳ��� ��ȸ"
End Sub


'## ���� ������.
Private Sub cboKaeyol_Click()
    If Me.Tag = "LOAD" Then Exit Sub
    
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01", "03"         '<< �ι�
            
            optTamgu(0).Caption = "����/����":      optTamgu(0).Tag = "00"
            optTamgu(1).Caption = "����":           optTamgu(1).Tag = "01"
            optTamgu(2).Caption = "����":           optTamgu(2).Tag = "02"
            optTamgu(3).Caption = "����":           optTamgu(3).Tag = "03"
            optTamgu(4).Caption = "�ѱ�":           optTamgu(4).Tag = "04"
            optTamgu(5).Caption = "�����":         optTamgu(5).Tag = "05"
            optTamgu(6).Caption = "����":           optTamgu(6).Tag = "06"
            optTamgu(7).Caption = "����":           optTamgu(7).Tag = "07"
            optTamgu(8).Caption = "��ġ":           optTamgu(8).Tag = "08"
            optTamgu(9).Caption = "�繮":           optTamgu(9).Tag = "09":             optTamgu(9).Visible = True:     optTamgu(9).BackColor = &HC000&
            optTamgu(10).Caption = "����":          optTamgu(10).Tag = "10":            optTamgu(10).Visible = True:    optTamgu(10).BackColor = &HFF&
            optTamgu(11).Caption = "����":          optTamgu(11).Tag = "11":            optTamgu(11).Visible = True:    optTamgu(11).BackColor = &HC0C0C0
            
            optTamgu(0).Value = True            '�⺻����
            
        Case "02"       '<< �ڿ�
            
            optTamgu(0).Caption = "����/����":      optTamgu(0).Tag = "00"
            optTamgu(1).Caption = "��1":            optTamgu(1).Tag = "51"
            optTamgu(2).Caption = "ȭ1":            optTamgu(2).Tag = "52"
            optTamgu(3).Caption = "��1":            optTamgu(3).Tag = "53"
            optTamgu(4).Caption = "��1":            optTamgu(4).Tag = "54"
            optTamgu(5).Caption = "��2":            optTamgu(5).Tag = "55"
            optTamgu(6).Caption = "ȭ2":            optTamgu(6).Tag = "56"
            optTamgu(7).Caption = "��2":            optTamgu(7).Tag = "57"
            optTamgu(8).Caption = "��2":            optTamgu(8).Tag = "58"
            
            optTamgu(9).Caption = "":               optTamgu(9).Tag = "00":             optTamgu(9).Visible = False:    optTamgu(9).BackColor = basModule.WhiteColor
            optTamgu(10).Caption = "":              optTamgu(10).Tag = "00":            optTamgu(10).Visible = False:   optTamgu(10).BackColor = basModule.WhiteColor
            optTamgu(11).Caption = "":              optTamgu(11).Tag = "00":            optTamgu(11).Visible = False:   optTamgu(11).BackColor = basModule.WhiteColor

            optTamgu(0).Value = True            '�⺻����
            
    End Select
    
    Call Find_LsnCD         '< �� ��ȸ
    
End Sub


'// ������
Private Sub sprGwamok_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim ni          As Integer
    Dim ninDex      As Integer
    Dim sTmp        As String

    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub

    With sprGwamok      '<< COLUMN���� ������.
        If .MaxCols = 0 Then Exit Sub
        
        For ni = 0 To optTamgu.UBound Step 1
            If optTamgu(ni).Value = True Then
                ninDex = ni
                Exit For
            End If
        Next ni
        
        If optTamgu(ninDex).Value = True Then
            .Row = Row:     .Row2 = Row
            .Col = Col:     .Col2 = Col
            .BlockMode = True
                .BackColor = optTamgu(ninDex).BackColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            Select Case optTamgu(ninDex).Tag
                Case "00"
                    .Text = ""
                Case Else
                    sTmp = optTamgu(ninDex).Caption
                    Call basFunction.Set_SprType_Text(sprGwamok, "center", "center", basFunction.LenKor(sTmp), sTmp)
            End Select
        End If
        
    End With
End Sub

Private Sub sprGwamok_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case vbRightButton
            With sprGwamok
                .Row = .ActiveRow
                .Col = .ActiveCol
                
                    .Text = ""
                
                .Row2 = .Row
                .Col2 = .Col
                .BlockMode = True
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
            End With
    End Select
        
End Sub

Private Sub sprGwamok_GotFocus()
    Dim ninDex  As Integer
    
    With sprGwamok
        If .ActiveCol < 1 Then Exit Sub
        If .ActiveRow < 1 Then Exit Sub
        
        
        .Row = .ActiveRow
        .Col = .ActiveCol
         
        
        Select Case Trim(.Text)
            Case "����":     ninDex = 1
            Case "����":     ninDex = 2
            Case "����":     ninDex = 3
            Case "�ѱ�":     ninDex = 4
            Case "�����", "����":   ninDex = 5
            Case "����":     ninDex = 6
            Case "����":     ninDex = 7
            Case "��ġ":     ninDex = 8
            Case "�繮":     ninDex = 9
            Case "����":     ninDex = 10
            Case "����":     ninDex = 11
            
            Case "��1":     ninDex = 1
            Case "ȭ1":     ninDex = 2
            Case "��1":     ninDex = 3
            Case "��1":     ninDex = 4
            Case "��2":     ninDex = 5
            Case "ȭ2":     ninDex = 6
            Case "��2":     ninDex = 7
            Case "��2":     ninDex = 8
            
            Case Else:      ninDex = 0
        End Select
        
        If ninDex = 0 Then
            .Row2 = .Row
            .Col2 = .Col
            .BlockMode = True
                .BackColor = basModule.WhiteColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        Else
            .Row2 = .Row
            .Col2 = .Col
            .BlockMode = True
                .BackColor = optTamgu(ninDex).BackColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        End If
    End With
End Sub

Private Sub sprGwamok_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ninDex  As Integer
    
    With sprGwamok
        If .ActiveCol < 1 Then Exit Sub
        If .ActiveRow < 1 Then Exit Sub
        
        
        .Row = .ActiveRow
        .Col = .ActiveCol
         
        
        Select Case Trim(.Text)
            Case "����":     ninDex = 1
            Case "����":     ninDex = 2
            Case "����":     ninDex = 3
            Case "�ѱ�":     ninDex = 4
            Case "�����", "����":   ninDex = 5
            Case "����":     ninDex = 6
            Case "����":     ninDex = 7
            Case "��ġ":     ninDex = 8
            Case "�繮":     ninDex = 9
            Case "����":     ninDex = 10
            Case "����":     ninDex = 11
            
            Case "��1":     ninDex = 1
            Case "ȭ1":     ninDex = 2
            Case "��1":     ninDex = 3
            Case "��1":     ninDex = 4
            Case "��2":     ninDex = 5
            Case "ȭ2":     ninDex = 6
            Case "��2":     ninDex = 7
            Case "��2":     ninDex = 8
            
            Case Else:      ninDex = 0
        End Select
        
        If ninDex = 0 Then
            .Row2 = .Row
            .Col2 = .Col
            .BlockMode = True
                .BackColor = basModule.WhiteColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        Else
            .Row2 = .Row
            .Col2 = .Col
            .BlockMode = True
                .BackColor = optTamgu(ninDex).BackColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        End If
    End With
End Sub

Private Sub sprGwamok_LostFocus()
    Dim ninDex  As Integer
    
    With sprGwamok
        If .ActiveCol < 1 Then Exit Sub
        If .ActiveRow < 1 Then Exit Sub
        
        
        .Row = .ActiveRow
        .Col = .ActiveCol
         
        
        Select Case Trim(.Text)
            Case "����":     ninDex = 1
            Case "����":     ninDex = 2
            Case "����":     ninDex = 3
            Case "�ѱ�":     ninDex = 4
            Case "�����", "����":   ninDex = 5
            Case "����":     ninDex = 6
            Case "����":     ninDex = 7
            Case "��ġ":     ninDex = 8
            Case "�繮":     ninDex = 9
            Case "����":     ninDex = 10
            Case "����":     ninDex = 11
            
            Case "��1":     ninDex = 1
            Case "ȭ1":     ninDex = 2
            Case "��1":     ninDex = 3
            Case "��1":     ninDex = 4
            Case "��2":     ninDex = 5
            Case "ȭ2":     ninDex = 6
            Case "��2":     ninDex = 7
            Case "��2":     ninDex = 8
            
            Case Else:      ninDex = 0
        End Select
        
        If ninDex = 0 Then
            .Row2 = .Row
            .Col2 = .Col
            .BlockMode = True
                .BackColor = basModule.WhiteColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        Else
            .Row2 = .Row
            .Col2 = .Col
            .BlockMode = True
                .BackColor = optTamgu(ninDex).BackColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        End If
    End With
End Sub

Private Sub sprGwamok_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ninDex  As Integer
    
    With sprGwamok
        If .ActiveCol < 1 Then Exit Sub
        If .ActiveRow < 1 Then Exit Sub
        
        .Row = .ActiveRow
        .Col = .ActiveCol
        
        If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
            .Text = ""
            .Row2 = .Row
            .Col2 = .Col
            
            .BlockMode = True
                .BackColor = basModule.WhiteColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            Exit Sub
        End If
        
        Select Case Trim(.Text)
            Case "����":     ninDex = 1
            Case "����":     ninDex = 2
            Case "����":     ninDex = 3
            Case "�ѱ�":     ninDex = 4
            Case "�����", "����":   ninDex = 5
            Case "����":     ninDex = 6
            Case "����":     ninDex = 7
            Case "��ġ":     ninDex = 8
            Case "�繮":     ninDex = 9
            Case "����":     ninDex = 10
            Case "����":     ninDex = 11
            
            Case "��1":     ninDex = 1
            Case "ȭ1":     ninDex = 2
            Case "��1":     ninDex = 3
            Case "��1":     ninDex = 4
            Case "��2":     ninDex = 5
            Case "ȭ2":     ninDex = 6
            Case "��2":     ninDex = 7
            Case "��2":     ninDex = 8
            
            Case Else:      ninDex = 0
        End Select
        
        If ninDex = 0 Then
            .Row2 = .Row
            .Col2 = .Col
            .BlockMode = True
                .BackColor = basModule.WhiteColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        Else
            .Row2 = .Row
            .Col2 = .Col
            .BlockMode = True
                .BackColor = optTamgu(ninDex).BackColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        End If
    End With
End Sub



'## ���񳻿� ���
Private Sub cmdSave_Click()
    Dim sTmp        As String
    
    cmdSave.Enabled = False
    
        With sprGwamok
            If .MaxCols = 0 Then
                MsgBox "����� ������ �����ϴ�.", vbExclamation + vbOKOnly, "������"
                cmdSave.Enabled = True
                Exit Sub
            End If
            
            sTmp = ""
            sTmp = "�� " & Trim(Left(cboKaeyol.Text, 30)) & " ���迭 "
            sTmp = sTmp & "�� " & Trim(Left(cboLsnType.Text, 30))
            sTmp = sTmp & " ��Ÿ������ �� ���ð��� ������ ����Ͻðڽ��ϱ�?"
            If MsgBox(sTmp, vbQuestion + vbYesNo, "���ð��� ���") = vbNo Then
                cmdSave.Enabled = True
                Exit Sub
            End If
            
            Call Save_inPutData
        
        End With
        
    
    cmdSave.Enabled = True
    
End Sub

Private Sub Save_inPutData()
    
    Dim DBCmd       As ADODB.Command        '<< �л� �� ���� ����ϱ�
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim nTot        As Long
    Dim nExeTot     As Long
    Dim nExe        As Long
    Dim nLength     As Long
    
    Dim nRow        As Long
    Dim nCol        As Integer
    Dim ni          As Integer
    
    Dim sTmp        As String
    Dim nTmp        As Long
    Dim sGwamok     As String
    
'>> ��Ϲ�� : ������ ��ϵ� type �� �ش��ϴ� ������ ��� ���� �� ó����.
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans

    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter

    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection


    '<< TYPE �� �ش��ϴ� ������ ��� ���� >>
    sStr = ""
    sStr = sStr & " DELETE "
    sStr = sStr & "   FROM SDLSN06TB "
    sStr = sStr & "  WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND KAEYOL  = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "    AND LSNTYPE = '" & Trim(Right(cboLsnType.Text, 30)) & "'"
    
'    '>> ACID
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("SEL_CLASS", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam

    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    nExe = 0
    DBCmd.Execute nExe, , -1
    
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    
    '<< ���� ���� ��� ���� >>
    With sprGwamok
        nTot = 0
        nExeTot = 0
        nExe = 0
        
        For nCol = 1 To .MaxCols Step 1
            For nRow = 1 To .MaxRows Step 1
                nTot = nTot + 1
                
                
                sStr = ""
                sStr = sStr & " INSERT INTO SDLSN06TB ( ACID       , KAEYOL     , LSNTYPE    , LSNCD      , ORD        , SUBJCD )"
                sStr = sStr & " VALUES ( "
                sStr = sStr & "       '" & Trim(basModule.SchCD) & "', "                '< ACID
                sStr = sStr & "       '" & Trim(Right(cboKaeyol.Text, 30)) & "', "      '< KAEYOL
                sStr = sStr & "       '" & Trim(Right(cboLsnType.Text, 30)) & "', "     '< LSNTYPE
                
                .Row = SpreadHeader
                .Col = nCol
                    sTmp = Trim(.Text)
                        sStr = sStr & "   '" & sTmp & "', "                             '< LSNCD
                .Row = nRow
                    sTmp = Trim(CStr(nRow))
                        sStr = sStr & "    " & sTmp & ", "                              '< ORD
                .Row = nRow
                .Col = nCol
                    Select Case Trim(.Text)                         '< ����üũ
                        Case "����":     sGwamok = "01"
                        Case "����":     sGwamok = "02"
                        Case "����":     sGwamok = "03"
                        Case "�ѱ�":     sGwamok = "04"
                        Case "�����", "����":     sGwamok = "05"
                        Case "����":     sGwamok = "06"
                        Case "����":     sGwamok = "07"
                        Case "��ġ":     sGwamok = "08"
                        Case "�繮":     sGwamok = "09"
                        Case "����":     sGwamok = "10"
                        Case "����":     sGwamok = "11"
                        
                        Case "��1":     sGwamok = "51"
                        Case "ȭ1":     sGwamok = "52"
                        Case "��1":     sGwamok = "53"
                        Case "��1":     sGwamok = "54"
                        Case "��2":     sGwamok = "55"
                        Case "ȭ2":     sGwamok = "56"
                        Case "��2":     sGwamok = "57"
                        Case "��2":     sGwamok = "58"
                        Case "":     sGwamok = ""
                    End Select
                    
                    sStr = sStr & "    '" & sGwamok & "' "                              '< SUBJCD
                sStr = sStr & " )"
                
                
                DBCmd.CommandText = sStr
                DBCmd.CommandType = adCmdText
                DBCmd.CommandTimeout = 30
        
                nExe = 0
                DBCmd.Execute nExe, , -1
        
                Do While basDataBase.DBConn.State And adStateExecuting
                    DoEvents
                Loop
        
                If nExe = 1 Then
                    nExeTot = nExeTot + 1
                End If
            
            Next nRow
        Next nCol
        
    End With
    
    '>> ó������ �����ؾ� ��.
    If nTot = nExeTot Then
        basDataBase.DBConn.CommitTrans
        MsgBox "���� ����Ͽ����ϴ�.", vbInformation + vbOKOnly, "������"
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "��� �� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������"
    End If
    
    ' NO ERROR
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Sub
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    MsgBox "���� ��� �� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
           Trim(CStr(Err.Number)) & " " & Err.Description, vbCritical + vbOKOnly, "������"
    
    On Error GoTo 0
End Sub


































