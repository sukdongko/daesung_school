VERSION 5.00
Begin VB.MDIForm MDI001 
   BackColor       =   &H8000000C&
   Caption         =   "����ȭ��"
   ClientHeight    =   11370
   ClientLeft      =   2055
   ClientTop       =   2415
   ClientWidth     =   14580
   Icon            =   "MDI001.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu mnu100 
      Caption         =   "���л���"
      Begin VB.Menu mnuSTD010_N 
         Caption         =   "�л����(�뷮��)"
      End
      Begin VB.Menu mnuSTD010 
         Caption         =   "�л� ��� (����,����,����,����,�λ�)"
      End
      Begin VB.Menu mnuSTD011_N 
         Caption         =   "�л���ü ��ȸ(�뷮��)"
      End
      Begin VB.Menu mnuSTD011 
         Caption         =   "�л���ü ��ȸ"
      End
      Begin VB.Menu mnuSTD012_N 
         Caption         =   "�л���ü ��ȸ (������)(�뷮��,����)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSTD012 
         Caption         =   "�л���ü ��ȸ (������)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSTD092 
         Caption         =   "�л������ ��ȸ"
      End
      Begin VB.Menu mnu100_Line11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSTD200 
         Caption         =   "���� && ����Ŭ���� ��� �� ��ȸ"
      End
      Begin VB.Menu mnu100_Line03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuINT 
         Caption         =   "���п��� ���"
         Begin VB.Menu mnuINT010_N 
            Caption         =   "���� ���п��� ��� (�뷮��)"
         End
         Begin VB.Menu mnuINT010 
            Caption         =   "���� ���п��� ���"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuINT110 
            Caption         =   "���� ���п��� ��� (�λ�)"
         End
         Begin VB.Menu mnuINT112 
            Caption         =   "���� ���п��� ��� (����)"
         End
         Begin VB.Menu mnuINT113 
            Caption         =   "���� ���п��� ��� (����,����)"
         End
         Begin VB.Menu mnuINT111 
            Caption         =   "���� ���п��� ��� (����)"
         End
         Begin VB.Menu mnu100_Line12 
            Caption         =   "-"
         End
         Begin VB.Menu mnuINT011 
            Caption         =   "���� ���п��� ��� (�뷮��, ����)"
         End
         Begin VB.Menu mnuINT012 
            Caption         =   "���� ���п��� ���"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuINT013 
            Caption         =   "���� ���п��� ��� (���̸�)"
         End
         Begin VB.Menu mnuINT014 
            Caption         =   "���� ���п��� ��� (����)"
         End
         Begin VB.Menu mnu100_Line05 
            Caption         =   "-"
         End
         Begin VB.Menu mnuINT020 
            Caption         =   "���� ���ͽ��� ���"
         End
         Begin VB.Menu mnuINT021 
            Caption         =   "���� ���ͽ��� ���"
         End
         Begin VB.Menu mnuINT022 
            Caption         =   "�뷮�� ���ͽ��� ���"
         End
         Begin VB.Menu mnuINT023 
            Caption         =   "���� ���ͽ��� ���"
         End
         Begin VB.Menu mnuINT024 
            Caption         =   "�λ� ���ͽ��� ���"
         End
         Begin VB.Menu mnu100_Line06 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMAT010 
            Caption         =   "���� ���� Ŭ���� ��� (�뷮��, ����)"
         End
         Begin VB.Menu mnuMAT010_J 
            Caption         =   "���� ���� Ŭ���� ��� (����)"
         End
      End
      Begin VB.Menu mnu100_Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSTD020 
         Caption         =   "�հݻ� ���"
      End
      Begin VB.Menu mnuSTD030 
         Caption         =   "��ϱ� �� ������� �ο�"
      End
      Begin VB.Menu mnu100_Line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSTD040 
         Caption         =   "�հݻ� �� �ð�ǥ �۾����� ���"
      End
      Begin VB.Menu mnu100_Line04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSTD050 
         Caption         =   "������ �л� ������ �����ϱ�"
      End
   End
   Begin VB.Menu mnu150 
      Caption         =   "�� �����ϱ�"
      Begin VB.Menu mnuLSN001 
         Caption         =   "�� ���"
      End
      Begin VB.Menu mnu150_Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLSN100 
         Caption         =   "�� �����ϱ�"
      End
   End
   Begin VB.Menu mnu200 
      Caption         =   "�̵��ð�ǥ �����"
      Begin VB.Menu mnuLSN001_CP 
         Caption         =   "�� ���� ���"
      End
      Begin VB.Menu mnuLSN002 
         Caption         =   "�̵� �� ���� ���"
      End
      Begin VB.Menu mnu200_Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLSN100_CP 
         Caption         =   "�� �����ϱ�"
      End
      Begin VB.Menu mnu200_Line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTMR020 
         Caption         =   "�̵����� �ð�ǥ ���_OLD"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTMR025 
         Caption         =   "�̵����� �ð�ǥ ���"
      End
      Begin VB.Menu mnuTMR027 
         Caption         =   "�̵����� ���񳻿� ���"
      End
   End
   Begin VB.Menu mnu250 
      Caption         =   "�ð�ǥ �����"
      Begin VB.Menu mnuLSN001_CP2 
         Caption         =   "�� ���� ���"
      End
      Begin VB.Menu mnu250_Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMTX011 
         Caption         =   "������ �ð�ǥ ���"
      End
      Begin VB.Menu mnu250_Line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTMR011 
         Caption         =   "���� �� ���纰 ����ֱ�"
      End
      Begin VB.Menu mnuTMR012 
         Caption         =   "���� �ü��ֱ�"
      End
      Begin VB.Menu mnuTMR015 
         Caption         =   "���� ���ǺҰ� �ð����"
      End
      Begin VB.Menu mnu250_Line03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTMR050 
         Caption         =   "��ü�ð�ǥ ����"
      End
   End
   Begin VB.Menu mnu300 
      Caption         =   "�ð�ǥ ���"
      Begin VB.Menu mnuPRT011 
         Caption         =   "�ݺ� �ð�ǥ ��� (�뷮��)"
      End
      Begin VB.Menu mnuPRT021 
         Caption         =   "���纰 �ð�ǥ ��� (�뷮��)"
      End
      Begin VB.Menu mnu300_Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPRT010 
         Caption         =   "�ݺ� �ð�ǥ ���"
      End
      Begin VB.Menu mnuPRT020 
         Caption         =   "���纰 �ð�ǥ ���"
      End
      Begin VB.Menu mnu300_Line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPRT050 
         Caption         =   "�� �ð�ǥ ��� (�뷮��)"
      End
      Begin VB.Menu mnu300_Line03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTMR060 
         Caption         =   "���� �⼮��"
      End
      Begin VB.Menu mnuPRT030 
         Caption         =   "���� �� �ݺ� �ð�ǥ �������� ��ȸ"
      End
   End
   Begin VB.Menu mnuEXM 
      Caption         =   "�������"
      Begin VB.Menu EXM010 
         Caption         =   "�л��������"
      End
      Begin VB.Menu EXM020 
         Caption         =   "�л��������"
      End
   End
   Begin VB.Menu TEST 
      Caption         =   "TEST"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MDI001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : MDI001
'   �� ��  �� �� : ����ȭ��
'
'   ��   ��   �� : 2007/08/22
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit


'�ǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢ�

Private Sub MenuYN()
    
    mnu100.Visible = False
    mnu150.Visible = False
    mnu200.Visible = False
    mnu250.Visible = False
    mnu300.Visible = False
    mnuEXM.Visible = False
    
    '�뷮���� �޴��� �����Ѵ�. �ٸ� �п����� ��� �޴� ����.
    
    If basModule.SchCD = "N" Then
        Select Case basModule.RegID
            Case "10000", "00002", "10003", "00001" '�迵������ (10000), �뷮��(00002), �躴ö(10003), ADMIN(00001)
                mnu100.Visible = True
                mnu150.Visible = True
                mnu200.Visible = True
                mnu250.Visible = True
                mnu300.Visible = True
                mnuEXM.Visible = True
                
            Case "10001"                            '������
                mnu100.Visible = True
                    mnuSTD010.Visible = True
                    mnuSTD011.Visible = True
                    mnuSTD012.Visible = True
                    mnuSTD092.Visible = True
                    
                    mnu100_Line11.Visible = False
                    mnuSTD200.Visible = False
                    mnu100_Line03.Visible = False
                    mnuINT.Visible = False
                    mnuSTD020.Visible = False
                    mnuSTD030.Visible = False
                    mnu100_Line02.Visible = False
                    mnuSTD040.Visible = False
                    mnu100_Line04.Visible = False
                    mnuSTD050.Visible = False
                    
                mnu150.Visible = False
                mnu200.Visible = False
                mnu250.Visible = False
                mnu300.Visible = False
                mnuEXM.Visible = True
                
                
            Case "10002"                            '������
                mnu100.Visible = True
                    mnuSTD010.Visible = True
                    mnuSTD011.Visible = True
                    mnuSTD012.Visible = True
                    mnuSTD092.Visible = True
                    
                    mnu100_Line11.Visible = False
                    mnuSTD200.Visible = False
                    mnu100_Line03.Visible = False
                    mnuINT.Visible = False
                    mnuSTD020.Visible = False
                    mnuSTD030.Visible = False
                    mnu100_Line02.Visible = False
                    mnuSTD040.Visible = False
                    mnu100_Line04.Visible = False
                    mnuSTD050.Visible = False
                
                mnu150.Visible = False
                mnu200.Visible = False
                mnu250.Visible = False
                mnu300.Visible = False
                mnuEXM.Visible = False
                
        End Select
    Else
        mnu100.Visible = True
        mnu150.Visible = True
        mnu200.Visible = True
        mnu250.Visible = True
        mnu300.Visible = True
        mnuEXM.Visible = True
        
    End If
    
End Sub

'�ǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢǢ�



Private Sub MDIForm_Load()
    Dim sAcID               As String
    Dim sConnections        As String
    
    Select Case Trim(basModule.SchCD)
        Case "N"
            sAcID = "�뷮��"
        Case "K"
            sAcID = "����"
        Case "S"
            sAcID = "����"
        Case "P"
            sAcID = "����mimac"
        Case "M"
            sAcID = "����mimac"
        
        Case "W"
            sAcID = "�ָ����Ǵ�"
        Case "Q"
            sAcID = "�߰����Ǵ�"
            
        Case "J"
            sAcID = "����"
        
        Case "B"
            sAcID = "�λ�"
        
    End Select
    
    
    Select Case UCase(Trim(basModule.connDB))
        Case "MIMAC"
            sConnections = "�Ǽ���"
        Case "DEV"
            sConnections = "���߿�"
        Case Else
            sConnections = "�Ǽ���"
    End Select
    
    '>> �ۼ�����üũ
    MDI001.Caption = "���л���. �ݹ��� �� �ð�ǥ �ۼ� ���α׷� (multiple) - 10.11.19 pm 04:02 " & "�� " & sConnections & " ��" & "�� " & sAcID & " ��  " & App.ProductName & " - ver " & App.Major & "." & App.Minor & "." & App.Revision
    
    Call MenuYN
    
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
    Unload TMR021
    Unload TMR022
    Unload INT900
    Unload TMR028
    
End Sub























Private Sub mnuINT010_N_Click()
    Load INT110_N
    INT110_N.Show
    INT110_N.ZOrder 0
End Sub



Private Sub mnuINT112_Click()
    Load INT112
    INT112.Show
    INT112.ZOrder 0
End Sub

Private Sub mnuINT113_Click()
    Load INT113
    INT113.Show
    INT113.ZOrder 0
End Sub



'>> ���л��� �л����
Private Sub mnuSTD010_Click()
    Load STD010
    STD010.Show
    STD010.ZOrder 0
    
End Sub

Private Sub mnuSTD010_N_Click()
    Load STD010_N
    STD010_N.Show
    STD010_N.ZOrder 0
End Sub

'>> �л���ü ��ȸ
Private Sub mnuSTD011_Click()
    Load STD011
    STD011.Show
    STD011.ZOrder 0
    
End Sub

Private Sub mnuSTD011_N_Click()
    Load STD011_N
    STD011_N.Show
    STD011_N.ZOrder 0
End Sub

Private Sub mnuSTD012_Click()
    Load STD012
    STD012.Show
    STD012.ZOrder 0
    
End Sub

Private Sub mnuSTD012_N_Click()
Load STD012_N
    STD012_N.Show
    STD012_N.ZOrder 0
End Sub

'>> ����л���ȸ
Private Sub mnuSTD092_Click()
    Load STD092
    STD092.Show
    STD092.ZOrder 0
    
End Sub


'>> ���п��� ��� // ���� ���п��� ���
Private Sub mnuINT011_Click()
    Load INT011
    INT011.Show
    INT011.ZOrder 0
    
End Sub

Private Sub mnuINT013_Click()
    Load INT013
    INT013.Show
    INT013.ZOrder 0
    
End Sub
Private Sub mnuINT014_Click()
    Load INT014
    INT014.Show
    INT014.ZOrder 0
End Sub



Private Sub mnuINT110_Click()
    Load INT110
    INT110.Show
    INT110.ZOrder 0
    
End Sub

Private Sub mnuINT111_Click()
    Load INT111
    INT111.Show
    INT111.ZOrder 0
    
End Sub

'>> ���п��� ��� // ���� ���ͽ���
Private Sub mnuINT020_Click()
    Load INT020
    INT020.Show
    INT020.ZOrder 0
    
End Sub

'>> ���п��� ��� // ���� ���ͽ���
Private Sub mnuINT021_Click()
    Load INT021
    INT021.Show
    INT021.ZOrder 0
    
End Sub

'>> ���п��� ��� // �뷮�� ���ͽ���
Private Sub mnuINT022_Click()
    Load INT022
    INT022.Show
    INT022.ZOrder 0
    
End Sub

'>> ���п��� ��� // ���� ���ͽ���
Private Sub mnuINT023_Click()
    Load INT023
    INT023.Show
    INT023.ZOrder 0

End Sub

'>> ���п��� ��� // �λ� ���ͽ���
Private Sub mnuINT024_Click()
    Load INT024
    INT024.Show
    INT024.ZOrder 0

End Sub



'>> ���� ���� Ŭ���� ���п���
Private Sub mnuMAT010_Click()
    Load MAT010
    MAT010.Show
    MAT010.ZOrder 0
    
End Sub

'>> ���� ���� Ŭ���� ���п��� ����
Private Sub mnuMAT010_J_Click()
    Load MAT011_J
    MAT011_J.Show
    MAT011_J.ZOrder 0
End Sub

'>> �հݻ� ���
Private Sub mnuSTD020_Click()
    Load STD020
    STD020.Show
    STD020.ZOrder 0
    
End Sub

'>> ��ϱ� �� ������� �ο�
Private Sub mnuSTD030_Click()
    Load STD031
    STD031.Show
    STD031.ZOrder 0
    
End Sub

'>> �����һ� ������ �����ϱ�
Private Sub mnuSTD090_Click()
    Load STD090
    STD090.Show
    STD090.ZOrder 0
    
End Sub



'>> ������ �л� ������ �����ϱ�
Private Sub mnuSTD050_Click()
    Load STD090
    STD090.Show
    STD090.ZOrder 0
    
End Sub

'>> �л� �ð�ǥ �۾�����...
Private Sub mnuSTD040_Click()
    Load STD040
    STD040.Show
    STD040.ZOrder 0
    
End Sub





'>> ������ ���
Private Sub mnuLSN001_Click()
    Load LSN001
    LSN001.Show
    LSN001.ZOrder 0
    
End Sub

'>> �� �����ϱ�
Private Sub mnuLSN100_Click()
    Load LSN100
    LSN100.Show
    LSN100.ZOrder 0
    
End Sub



'>> �� ���
Private Sub mnuLSN001_CP_Click()
    Load LSN001
    LSN001.Show
    LSN001.ZOrder 0
    
End Sub


'>> �̵��� ���� ���
Private Sub mnuLSN002_Click()
    Load LSN002
    LSN002.Show
    LSN002.ZOrder 0
End Sub

'>> �� �����ϱ�
Private Sub mnuLSN100_CP_Click()
    Load LSN100
    LSN100.Show
    LSN100.ZOrder 0
End Sub



'>> ������ �ð�ǥ �ڵ� ���
Private Sub mnuMTX011_Click()
    Load MTX011
    MTX011.Show
    MTX011.ZOrder 0
    
End Sub





'>> ���� �� ���纰 ����ֱ�
Private Sub mnuTMR011_Click()
    Load TMR011
    TMR011.Show
    TMR011.ZOrder 0
    
End Sub

'>> ���� �ü��ֱ�
Private Sub mnuTMR012_Click()
    Load TMR012
    TMR012.Show
    TMR012.ZOrder 0
    
End Sub

'>> ���� ���ǺҰ� �ð����
Private Sub mnuTMR015_Click()
    Load TMR015
    TMR015.Show
    TMR015.ZOrder 0
    
End Sub



'>> �̵������ð�ǥ ���
Private Sub mnuTMR025_Click()
    Load TMR026
    TMR026.Show
    TMR026.ZOrder 0
    
End Sub


Private Sub mnuTMR027_Click()
    Load TMR028
    TMR028.Show
    TMR028.ZOrder 0
    
End Sub

'## ��ü �ð�ǥ ���� : �ݺ�
Private Sub mnuTMR050_Click()
    Load TMR051
    TMR051.Show
    TMR051.ZOrder 0
    
End Sub


'>> �� ���� ���
Private Sub mnuLSN001_CP2_Click()
    Load LSN001
    LSN001.Show
    LSN001.ZOrder 0

End Sub


'>> �ݺ� �ð�ǥ ���
Private Sub mnuPRT010_Click()
    Load PRT010
    PRT010.Show
    PRT010.ZOrder 0
    
End Sub

'>> ���纰 �ð�ǥ ���
Private Sub mnuPRT020_Click()
    Load PRT020
    PRT020.Show
    PRT020.ZOrder 0
    
End Sub


'>> ���� �� �ݺ� �ð�ǥ �������� ��ȸ
Private Sub mnuPRT030_Click()
    Load PRT031
    PRT031.Show
    PRT031.ZOrder 0

End Sub


'## �뷮�� ��¾�� �����û : 2008.02.19
Private Sub mnuPRT011_Click()
'    Load PRT011
'    PRT011.Show
'    PRT011.ZOrder 0
    
    '>> 2008.02.25 ����
    Load PRT012
    PRT012.Show
    PRT012.ZOrder 0
    
End Sub

Private Sub mnuPRT021_Click()
'    Load PRT021
'    PRT021.Show
'    PRT021.ZOrder 0
    
    '>> 2008.02.25 ����
    Load PRT022
    PRT022.Show
    PRT022.ZOrder 0
    
End Sub

Private Sub mnuPRT050_Click()
'    Load PRT050
'    PRT050.Show
'    PRT050.ZOrder 0
    
    Load PRT051
    PRT051.Show
    PRT051.ZOrder 0
    
End Sub

'## ���� �⼮��
Private Sub mnuTMR060_Click()
    Load TMR060
    TMR060.Show
    TMR060.ZOrder 0
    
End Sub

'>> ��ϱ� �κ� �׽�Ʈ
Private Sub TEST_Click()
    Load TMR028
    TMR028.Show
    TMR028.ZOrder 0
    
End Sub




'>> ���� & ����Ŭ���� ��� �� ��ȸ
Private Sub mnuSTD200_Click()
    Load STD200
    STD200.Show
    STD200.ZOrder 0
End Sub

'>> �л��������
Private Sub EXM010_Click()
    Load EXM100
    EXM100.Show
    EXM100.ZOrder 0
    
End Sub

'>> �л�����
Private Sub EXM020_Click()
    Load EXM110
    EXM110.Show
    EXM110.ZOrder 0
    
End Sub
