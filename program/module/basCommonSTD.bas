Attribute VB_Name = "basCommonSTD"
    Option Explicit
    
    
    Dim g_


Function Set_Spread_Design1(ByRef sprControl As Object)

    With sprControl
        .ShadowColor = basModule.ShadowColor1
        .ShadowDark = basModule.ShadowDark1
        .ShadowText = basModule.ShadowText1
        .GridColor = basModule.GridColor1
        .GrayAreaBackColor = basModule.GrayAreaBackColor1
    End With
    
End Function

''�뷮�� �迭 �޺��ڽ� ����.
'Function Init_Kaeyol_N(ByRef cboControl As Object)
'    With cboControl
'        .Clear
'        .AddItem "�ι�" & Space(30) & "01"
'        .AddItem "�ڿ�" & Space(30) & "02"
'
'
'    '<< �迭 >> : 2008.01.09
'        If Trim(basModule.SchCD) = "N" Then             '< �뷮��
'
'            .AddItem "������ι�" & Space(30) & "21"
'            .AddItem "������ڿ�" & Space(30) & "22"
'            .AddItem "��ü" & Space(30) & "03"
'            .AddItem "����(��)" & Space(30) & "04"
'            .AddItem "�ι�����" & Space(30) & "05"
'            .AddItem "�ڿ�����" & Space(30) & "06"
'
'            .AddItem "�ι�-��" & Space(30) & "07"
'            .AddItem "�ڿ�-��" & Space(30) & "08"
'            '.AddItem "�����ι�-��" & Space(30) & "09"
'            '.AddItem "�����ڿ�-��" & Space(30) & "10"
'
'            .AddItem "��)�ι�" & Space(30) & "11"
'            .AddItem "��)�ڿ�" & Space(30) & "12"
'            .AddItem "��)��ü" & Space(30) & "13"
'            .AddItem "��)����(��)" & Space(30) & "14"
'            .AddItem "��)�ι�����" & Space(30) & "15"
'            .AddItem "��)�ڿ�����" & Space(30) & "16"
'
'        End If
'    '<< �迭 >> : 2008.01.10
'        'If Trim(basModule.SchCD) = "K" Then             '< ����
'        Select Case Trim(basModule.SchCD)
'            Case "K", "W", "Q"
'                .AddItem "�ָ�����" & Space(30) & "04"
'                .AddItem "�ָ��Ǵ�" & Space(30) & "05"
'
'                .AddItem "�߰�����" & Space(30) & "06"
'                .AddItem "�߰��Ǵ�" & Space(30) & "07"
'
'                .AddItem "�������ι�" & Space(30) & "11"
'                .AddItem "�������ڿ�" & Space(30) & "12"
'
'                .AddItem "�������ι�16" & Space(30) & "16"
'                .AddItem "�������ڿ�17" & Space(30) & "17"
'
'        End Select
'
'    '<< �迭 >> : 2008.02.15
'        Select Case Trim(basModule.SchCD)               '< ����
'            Case "S"
''                .AddItem "��ü��" & Space(30) & "03"
''
''                .AddItem "�ι�����" & Space(30) & "05"
''                .AddItem "�ڿ�����" & Space(30) & "06"
''
'                .AddItem "�ż��ι�" & Space(30) & "11"
'                .AddItem "�ż��ڿ�" & Space(30) & "12"
'
''                .AddItem "�ι������̾�" & Space(30) & "18"
''                .AddItem "�ڿ������̾�" & Space(30) & "19"
'
'                .AddItem "�����Ư���ι�" & Space(30) & "21"
'                .AddItem "�����Ư���ڿ�" & Space(30) & "22"
'
'                .AddItem "�߰�������ι�" & Space(30) & "21"
'                .AddItem "�߰�������ڿ�" & Space(30) & "22"
'
'        End Select
'
'        Select Case Trim(basModule.SchCD)               '< ����
'            Case "J"
'                .AddItem "�ż��ι�" & Space(30) & "11"
'                .AddItem "�ż��ڿ�" & Space(30) & "12"
'                .AddItem "�ι������̾�" & Space(30) & "18"
'                .AddItem "�ڿ������̾�" & Space(30) & "19"
'                .AddItem "�����Ư���ι�" & Space(30) & "21"
'                .AddItem "�����Ư���ڿ�" & Space(30) & "22"
'        End Select
'
'    '<< �迭 >> : 2009.01.09
'        If Trim(basModule.SchCD) = "B" Then             '< �λ�
'
'            .AddItem "�ι�PS��" & Space(30) & "23"
'            .AddItem "�ڿ�PM��" & Space(30) & "24"
'
'            .AddItem "���м����ι�" & Space(30) & "05"
'            .AddItem "���м����ڿ�" & Space(30) & "06"
'
'            .AddItem "��.����ι�" & Space(30) & "07"
'            .AddItem "��.����ڿ�" & Space(30) & "08"
'
'            .AddItem "��ȭ�ι�" & Space(30) & "09"
'            .AddItem "��ȭ�ڿ�" & Space(30) & "10"
'        End If
'
'        Select Case Trim(basModule.SchCD)               '< ����
'            Case "M"
'                .AddItem "�����Ư���ι�" & Space(30) & "21"
'                .AddItem "�����Ư���ڿ�" & Space(30) & "22"
'        End Select
'
'        .ListIndex = 0
'    End With
'End Function


Function Init_CboKaeyolDefault(ByRef cboControl As Object)
    With cboControl
        .Clear
        .AddItem "�ι�" & Space(30) & "01"
        .AddItem "�ڿ�" & Space(30) & "02"
        
        
    '<< �迭 >> : 2008.01.09
        If Trim(basModule.SchCD) = "N" Then             '< �뷮��
        
            .AddItem "������ι�" & Space(30) & "21"
            .AddItem "������ڿ�" & Space(30) & "22"
            .AddItem "��ü" & Space(30) & "03"
            .AddItem "����(��)" & Space(30) & "04"
            .AddItem "�ι�����" & Space(30) & "05"
            .AddItem "�ڿ�����" & Space(30) & "06"
            
            .AddItem "�ι�-��" & Space(30) & "07"
            .AddItem "�ڿ�-��" & Space(30) & "08"
            '.AddItem "�����ι�-��" & Space(30) & "09"
            '.AddItem "�����ڿ�-��" & Space(30) & "10"
            
            .AddItem "��)�ι�" & Space(30) & "11"
            .AddItem "��)�ڿ�" & Space(30) & "12"
            .AddItem "��)��ü" & Space(30) & "13"
            .AddItem "��)����(��)" & Space(30) & "14"
            .AddItem "��)�ι�����" & Space(30) & "15"
            .AddItem "��)�ڿ�����" & Space(30) & "16"
            
            
        End If
    '<< �迭 >> : 2008.01.10
        'If Trim(basModule.SchCD) = "K" Then             '< ����
        Select Case Trim(basModule.SchCD)
            Case "K", "W", "Q"
                .AddItem "�ָ�����" & Space(30) & "04"
                .AddItem "�ָ��Ǵ�" & Space(30) & "05"
                
                .AddItem "�߰�����" & Space(30) & "06"
                .AddItem "�߰��Ǵ�" & Space(30) & "07"
                
                .AddItem "�������ι�" & Space(30) & "11"
                .AddItem "�������ڿ�" & Space(30) & "12"
                
                .AddItem "�������ι�16" & Space(30) & "16"
                .AddItem "�������ڿ�17" & Space(30) & "17"
                
                .AddItem "���ſ�����ι�" & Space(30) & "19"
                .AddItem "���ſ�����ڿ�" & Space(30) & "20"
        End Select
    
        '<< �迭 >> : 2008.02.15
        Select Case Trim(basModule.SchCD)               '< ����
            Case "S"
'                .AddItem "��ü��" & Space(30) & "03"
'
'                .AddItem "�ι�����" & Space(30) & "05"
'                .AddItem "�ڿ�����" & Space(30) & "06"
'
                .AddItem "�ż��ι�" & Space(30) & "11"
                .AddItem "�ż��ڿ�" & Space(30) & "12"
                
'                .AddItem "�ι������̾�" & Space(30) & "18"
'                .AddItem "�ڿ������̾�" & Space(30) & "19"

                .AddItem "�����Ư���ι�" & Space(30) & "21"
                .AddItem "�����Ư���ڿ�" & Space(30) & "22"
                
                .AddItem "�߰�������ι�" & Space(30) & "21"
                .AddItem "�߰�������ڿ�" & Space(30) & "22"
                
        End Select
        
        
        Select Case Trim(basModule.SchCD)               '< ����
            Case "J"
                .AddItem "�ż��ι�" & Space(30) & "11"
                .AddItem "�ż��ڿ�" & Space(30) & "12"
                .AddItem "�ι������̾�" & Space(30) & "18"
                .AddItem "�ڿ������̾�" & Space(30) & "19"
                .AddItem "�����Ư���ι�" & Space(30) & "21"
                .AddItem "�����Ư���ڿ�" & Space(30) & "22"
        End Select
        
    '<< �迭 >> : 2009.01.09
        If Trim(basModule.SchCD) = "B" Then             '< �λ�
            
            .AddItem "�ι�PS��" & Space(30) & "23"
            .AddItem "�ڿ�PM��" & Space(30) & "24"
            
            .AddItem "�����ι�" & Space(30) & "05"
            .AddItem "�����ڿ�" & Space(30) & "06"
            
            .AddItem "��.����ι�" & Space(30) & "07"
            .AddItem "��.����ڿ�" & Space(30) & "08"
            
            .AddItem "��ȭ�ι�" & Space(30) & "09"
            .AddItem "��ȭ�ڿ�" & Space(30) & "10"
        End If
        
        Select Case Trim(basModule.SchCD)               '< ����
            Case "M"
                .AddItem "�����Ư���ι�" & Space(30) & "21"
                .AddItem "�����Ư���ڿ�" & Space(30) & "22"
        End Select
    
        .ListIndex = 0
    End With
End Function


Function Set_CboKaeyol(ByRef cboControl As Object, ByVal SchCD As String, ByVal kaeyol As String) As Boolean
    
    If IsNull(SchCD) = True Or SchCD = "" Then
        Set_CboKaeyol = False
        Exit Function
    End If
    
    If IsNull(kaeyol) = True Or kaeyol = "" Then
        Set_CboKaeyol = False
        Exit Function
    End If
    
    
    If Trim(SchCD) = "N" Then
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                Case "03"
                    cboControl.ListIndex = 4
                Case "04"
                    cboControl.ListIndex = 5
                Case "05"
                    cboControl.ListIndex = 6
                Case "06"
                    cboControl.ListIndex = 7
                    
                    
                Case "07"
                    cboControl.ListIndex = 8
                Case "08"
                    cboControl.ListIndex = 9
                'Case "09"
                '    cboControl.ListIndex = 8
                'Case "10"
                '    cboControl.ListIndex = 9
                    
                Case "11"
                    cboControl.ListIndex = 10
                Case "12"
                    cboControl.ListIndex = 11
                Case "13"
                    cboControl.ListIndex = 12
                Case "14"
                    cboControl.ListIndex = 13
                Case "15"
                    cboControl.ListIndex = 14
                Case "16"
                    cboControl.ListIndex = 15
                Case "21"
                    cboControl.ListIndex = 2
                Case "22"
                    cboControl.ListIndex = 3
            End Select
        End If
        
    ElseIf Trim(SchCD) = "B" Then
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                Case "05"
                    cboControl.ListIndex = 4
                Case "06"
                    cboControl.ListIndex = 5
                    
                Case "07"
                    cboControl.ListIndex = 6
                Case "08"
                    cboControl.ListIndex = 7
                Case "09"
                    cboControl.ListIndex = 8
                Case "10"
                    cboControl.ListIndex = 9
                Case "23"
                    cboControl.ListIndex = 2
                Case "24"
                    cboControl.ListIndex = 3
            End Select
        End If
        
    ElseIf (Trim(SchCD) = "K") Or (Trim(SchCD) = "W") Or (Trim(SchCD) = "Q") Then
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                    
                Case "04"
                    cboControl.ListIndex = 2
                Case "05"
                    cboControl.ListIndex = 3
                Case "06"
                    cboControl.ListIndex = 4
                Case "07"
                    cboControl.ListIndex = 5
                    
                Case "11"
                    cboControl.ListIndex = 6
                Case "12"
                    cboControl.ListIndex = 7
                    
                Case "16"
                    cboControl.ListIndex = 8
                Case "17"
                    cboControl.ListIndex = 9
                    
                Case "19"
                    cboControl.ListIndex = 10
                Case "20"
                    cboControl.ListIndex = 11
                    
            End Select
        End If
        
    ElseIf Trim(SchCD) = "S" Then
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                    
        ''                        Case "03"
        ''                            cboControl.ListIndex = 2
        ''
        ''                        Case "05"
        ''                            cboControl.ListIndex = 3
        ''                        Case "06"
        ''                            cboControl.ListIndex = 4
        ''
        ''                        Case "11"
        ''                            cboControl.ListIndex = 5
        ''                        Case "12"
        ''                            cboControl.ListIndex = 6
        '
        '                        Case "18"
        '                            cboControl.ListIndex = 7
        '                        Case "19"
        '                            cboControl.ListIndex = 8
                    
        '                        Case "18"
        '                            cboControl.ListIndex = 2
        '                        Case "19"
        '                            cboControl.ListIndex = 3
                Case "21"
                    cboControl.ListIndex = 2
                Case "22"
                    cboControl.ListIndex = 3
                    
                    
            End Select
        End If
    ElseIf Trim(SchCD) = "P" Then                 '< ����
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                    
                Case "03"
                    cboControl.ListIndex = 2
                Case "04"
                    cboControl.ListIndex = 3
            End Select
        End If
        
    ElseIf Trim(SchCD) = "J" Then                 '< ����
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                    
                Case "11"
                    cboControl.ListIndex = 2
                Case "12"
                    cboControl.ListIndex = 3
                    
                Case "18"
                    cboControl.ListIndex = 4
                Case "19"
                    cboControl.ListIndex = 5
                    
            End Select
        End If
    ElseIf Trim(SchCD) = "M" Then                 '< ����
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                    
                Case "21"
                    cboControl.ListIndex = 2
                Case "22"
                    cboControl.ListIndex = 3
            End Select
        End If
        
    Else
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                Case "03"
                    cboControl.ListIndex = 2
            End Select
        End If
    End If
End Function
'�п�
Function Init_CboSch(ByRef cboControl As Object)
    With cboControl
        .Clear
        .AddItem "����" & Space(30) & "X"
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
End Function

'�п�
Function Set_CboSch(ByRef cboControl As Object, ByVal sSch As String)

    Select Case Trim(sSch)
        Case "N"
            cboControl.ListIndex = 1
        Case "K"
            cboControl.ListIndex = 2
        Case "S"
            cboControl.ListIndex = 3
        Case "P"
            cboControl.ListIndex = 4
        Case "M"
            cboControl.ListIndex = 5
        Case "W"
            cboControl.ListIndex = 6
        Case "Q"
            cboControl.ListIndex = 7
        Case "J"
            cboControl.ListIndex = 8
        Case "B"
            cboControl.ListIndex = 9
        Case Else
            cboControl.ListIndex = 0
    End Select
End Function


'�հ�
Function Init_PassCN(ByRef cboControl As Object)
    With cboControl
        .Clear
        .AddItem "��ü" & Space(30) & "ALL"
        .AddItem "1��" & Space(30) & "1"
        .AddItem "2��" & Space(30) & "2"
        .AddItem "3��" & Space(30) & "3"
        .AddItem "4��" & Space(30) & "4"
        
        .ListIndex = 0
    End With
End Function

'����
Function Init_Pay(ByRef cboControl As Object)
    With cboControl
        .Clear
        .AddItem "��ü" & Space(30) & "ALL"
        .AddItem "����" & Space(30) & "OK"
        .AddItem "�̰���" & Space(30) & "NOT"
        
        .ListIndex = 0
    End With
End Function

'��������
Function Init_ExmType(ByRef cboControl As Object)
    With cboControl
        .Clear
        .AddItem "��ü" & Space(30) & "ALL"
        .AddItem "������" & Space(30) & "1"
        .AddItem "������" & Space(30) & "0"
        
        .ListIndex = 0
    End With
End Function

'���ͳ�/�п�
Function Init_InGbn(ByRef cboControl As Object)
    With cboControl
        .Clear
        .AddItem "��ü" & Space(30) & "ALL"
        .AddItem "���ͳ�" & Space(30) & "INT"
        .AddItem "�п�" & Space(30) & "HAK"
        
        .ListIndex = 0
    End With
End Function

'���
Function Init_Mu_type(ByRef cboControl As Object)
    With cboControl
        .Clear
        
        .AddItem "���ɵ��" & Space(30) & "1"   '����
        .AddItem "2013 6�� �򰡿�" & Space(30) & "2"
        .AddItem "2013 9�� �򰡿�" & Space(30) & "3"
        
        If basModule.SchCD = "N" Or basModule.SchCD = "S" _
            Or basModule.SchCD = "J" Or basModule.SchCD = "K" Or basModule.SchCD = "M" Then
            .AddItem "���ŵ��" & Space(30) & "9"
        End If
        
        .AddItem "����" & Space(30) & "X"
        
        .Enabled = True
        .ListIndex = .ListCount - 1
        
    End With
End Function

'�������� ����
Function Init_PTS_Sel(ByRef cboControl As Object)
    With cboControl
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem ""
        
        .Enabled = False
        .ListIndex = 2
    End With
End Function

'ī��
Function Init_Card(ByRef cboControl As Object)
    With cboControl
        Select Case Trim(basModule.SchCD)
            Case "N", "K", "W", "Q", "S"
                .AddItem "�Ƹ߽�ī��               AMX"
                .AddItem "��������ī��             CBB"
                .AddItem "���̳ʽ�ī��             DIN"
                .AddItem "�ѹ�����ī��             KAB"
                .AddItem "����ī��                 KWB"
                .AddItem "����ī��                 NLC"
                .AddItem "�ż���ī��               SIN"
                .AddItem "BCī��                   BCC"
                .AddItem "��������ī��             CJB"
                .AddItem "�ϳ�����ī��             HNB"
                .AddItem "��ȯ����ī��             KEB"
                .AddItem "LGī��                   LGC"
                .AddItem "��ȭ����ī��             PHB"
                .AddItem "�Ｚī��                 WIN"
                .AddItem "�ܱ�����ī��             BRD"
                .AddItem "��������ī��             CNB"
                .AddItem "JCBī��                  JCB"
                .AddItem "��������ī��             KJB"
                .AddItem "����ī��                 NFF"
                .AddItem "��������ī��             SHB"
                
            Case "M", "P", "J", "B"
    
                '20121221
                .AddItem "KB����ī��        CCKM"
                .AddItem "NHä��ī��        CCNH"
                .AddItem "�ż����ѹ�        CCSG"
                .AddItem "��Ƽī��          CCCT"
                .AddItem "�ѹ�ī��          CCHM"
                .AddItem "�ؿܺ���          CVSF"
                .AddItem "�����Ƹ߽�        CCAM"
                .AddItem "�Ե�ī��          CCLO"
                .AddItem "�ؿܾƸ߽�        CAMF"
                .AddItem "BCī��            CCBC"
                .AddItem "�츮ī��          CCPH"
                .AddItem "�ϳ�SKī��        CCHN"
                .AddItem "�Ｚī��          CCSS"
                .AddItem "����ī��          CCKJ"
                .AddItem "����ī��          CCSU"
                .AddItem "����ī��          CCCU"
                .AddItem "����ī��          CCSH"
                .AddItem "����ī��          CCJB"
                .AddItem "����ī��          CCCJ"
                .AddItem "����ī��          CCLG"
                .AddItem "�ؿܸ�����        CMCF"
                .AddItem "�ؿ�JCB           CJCF"
                .AddItem "��ȯī��          CCKE"
                .AddItem "����ī��          CCDI"
                .AddItem "����ī��          CCSB"
                .AddItem "����ī��          CCKD"
                .AddItem "����ī��          CCUF"
'                .AddItem "BCī��                      CCBC"
'                .AddItem "����ī��                    CCKM"
'                .AddItem "LGī��                      CCLG"
'                .AddItem "�Ｚī��                    CCSS"
'                .AddItem "��ȯī��                    CCKE"
'                .AddItem "����ī��                    CCSH"
'                .AddItem "����ī��                    CCSU"
'                .AddItem "��������                    CCKJ"
'                .AddItem "��������                    CCKW"
'                .AddItem "�ϳ�����                    CCHN"
'                .AddItem "�����Ƹ߽�                  CCAM"
'                .AddItem "�ؿܾƸ߽�                  CAMF"
'                .AddItem "�ѹ�����                    CCYJ"
'                .AddItem "����ī��                    CCCH"
'                .AddItem "��ȭ����                    CCPH"
'                .AddItem "��������                    CCCJ"
'                .AddItem "��������                    CCJB"
'                .AddItem "����ī��                    CCDI"
'                .AddItem "��Ƽ����                    CCCT"
'                .AddItem "��������                    CCDN"
'                .AddItem "�ؿܺ���                    CVSF"
'                .AddItem "�ؿܸ���Ÿī��              CMCF"
'                .AddItem "�ؿ�JCBī��                 CJCF"
'                .AddItem "�Ե�ī��                    CCLO"
                
        End Select
        .ListIndex = 0
    End With

End Function

'Ŭ���� �޺� �ʱ�ȭ
Sub Init_Clinic(ByRef chkClinic_L As Object, ByRef chkClinic_M As Object, ByRef chkClinic_E As Object)

    If SchCD = "N" Or SchCD = "S" Then
        Dim i As Integer
    
        For i = 0 To chkClinic_L.count - 1
            chkClinic_L(i).value = False
        Next i
        
        For i = 0 To chkClinic_M.count - 1
            chkClinic_M(i).value = False
        Next i
        
        For i = 0 To chkClinic_E.count - 1
            chkClinic_E(i).value = False
        Next i
        
    Else
        
        
    End If

End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'�ʱ�ȭ ��
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'������ ��������/�ҷ�����
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Function Get_SchName(sSch As String)

    If IsNull(sSch) = True Then
        Get_SchName = ""
        Exit Function
    End If
    
    Dim sTmp As String
    Select Case Trim(sSch)
        Case "N"
            sTmp = "�뷮��"
        Case "K"
            sTmp = "����"
        Case "S"
            sTmp = "����"
        Case "P"
            sTmp = "���� M"
        Case "M"
            sTmp = "���� M"
        Case "W"
            sTmp = "�ָ����Ǵ�"
        Case "Q"
            sTmp = "�߰����Ǵ�"
        Case "J"
            sTmp = "����"
        Case "B"
            sTmp = "�λ�"
        Case "E"
            sTmp = "�������(��õ)"
        Case Else
            sTmp = ""
    End Select
    
    Get_SchName = sTmp
End Function


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'������ ��������/�ҷ�����
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    
'Ŭ���� �޺� ����
Sub Set_Clinic(ByRef chkClinic_L As Object, ByRef chkClinic_M As Object, ByRef chkClinic_E As Object, ByVal SEL7 As String)

    Dim sDiv()      As String
    Dim ni          As Integer
    
    
    sDiv = Split(SEL7, "|", -1, vbTextCompare)
    
    For ni = 0 To UBound(sDiv) - 1 Step 1
        If sDiv(ni) <> "" Then
            If sDiv(ni) >= CLINIC_L_CLASS And sDiv(ni) < CLINIC_L_CLASS + 10 Then
            
                chkClinic_L(CInt(sDiv(ni)) - CLINIC_L_CLASS).value = True
                
            ElseIf sDiv(ni) >= CLINIC_M_CLASS And sDiv(ni) < CLINIC_M_CLASS + 10 Then
            
                chkClinic_M(CInt(sDiv(ni)) - CLINIC_M_CLASS).value = True
                
            ElseIf sDiv(ni) >= CLINIC_E_CLASS And sDiv(ni) < CLINIC_E_CLASS + 10 Then
            
                chkClinic_E(CInt(sDiv(ni)) - CLINIC_E_CLASS).value = True
                
            End If
        End If
    Next ni
    
End Sub

'���
Public Sub Set_Mu_type(ByRef cboControl As Object, ByVal val As Integer)
    Select Case val
        Case "1"
            cboControl.ListIndex = 0 '���ɵ��
        Case "2"
            cboControl.ListIndex = 1 '6�� ����
        Case "3"
            cboControl.ListIndex = 2 '9�� ����
        Case "9"
            cboControl.ListIndex = 3 '���ŵ��
    End Select
    
End Sub


'����� �� ��������
Public Function Get_StrMuType(ByVal value)
     Select Case value
        Case "1"
            Get_StrMuType = "���ɵ��"
        Case "2"
            Get_StrMuType = "6�� ����"
        Case "3"
            Get_StrMuType = "9�� ����"
        Case "9"
            Get_StrMuType = "���ŵ��"
    End Select
End Function

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'�п��� ��������
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Public Function Get_StrGongji() As String()
    Dim strReturn() As String
    
    '>> �г⺰ ����
    
    Select Case Trim(basModule.SchCD)
    
        Case "N"
            ReDim strReturn(3)
            strReturn(0) = "�� ����, ����, ����, ���� �� (��ȭ)�Ǵ�(�⺻) ���� 1������ �����ؾ� �ϸ�, Ž������ 4���� �� 1������ ������ �� �ֽ��ϴ�."
            strReturn(1) = "�� �ι���� ��Ȱ�� ����, ������ ���, ��������, ���ƽþƻ�, �����, ����, ��2�ܱ���, �ڿ���� ���Х�(4����)�� ��� ���Թݺ��� �����մϴ�."
            strReturn(2) = "�� �ݴ� ������ �� ������ ���� �й� �Ǵ� �չ��� �� �ֽ��ϴ�."
            strReturn(3) = "�� �ι����(����B, ����A, ����B) / �ڿ���(����A, ����B, ����B��)���� �����մϴ�."
           
        Case "K", "W", "Q"
            ReDim strReturn(2)
            strReturn(0) = "���ι��� ��ȸŽ�� �� ���� ��ġ, �����, ��������, ���ƽþƻ�, ������ ���, ��Ȱ�� ���� �� ��2�ܱ���� ���Թݿ��� �����Ͽ� ������ �� �ֽ��ϴ�."
            strReturn(1) = "���ڿ��� ����Ž�� �� ������, ȭ�Х�, ������Х�, �������Х��� ���Թݿ��� �����Ͽ� ������ �� �ֽ��ϴ�."
            strReturn(2) = "���ι���(����B, ����A, ����B��)/�ڿ���(����A, ����B, ����B��)���� �����մϴ�."
            
        Case "S"
            ReDim strReturn(3)
            strReturn(0) = "�� ����, ����, ����, ���� �� (��ȭ)�Ǵ�(�⺻) ���� 2������ �����ؾ� �ϸ�, Ž�������� 1������ ������ �� �ֽ��ϴ�."
            strReturn(1) = "�� �ι���� ��Ȱ�� ����, ������ ���, ��������, ���ƽþƻ�, �����, ����, ��2�ܱ���, �ڿ���� ���Х�(4����)�� ��� ���Թݺ��� �����մϴ�."
            strReturn(2) = "�� �ݴ� ������ �� ������ ���� �й� �Ǵ� �չ��� �� �ֽ��ϴ�."
            strReturn(3) = "�� �ι����(����B, ����A, ����B) / �ڿ���(����A, ����B, ����B��)���� �����մϴ�."
                  
        Case "P"
            ReDim strReturn(1)
            strReturn(0) = ""
            strReturn(1) = ""
                             
         Case "M"
            ReDim strReturn(2)
            strReturn(0) = "���ι��� ��ȸŽ�� �� �����, ��������, ���ƽþƻ�, ��Ȱ�� ���� �� ��2�ܱ���� ���Թݿ��� �����Ͽ� ������ �� �ֽ��ϴ�."
            strReturn(1) = "���ι���(����B, ����A, ����B��)/�ڿ���(����A, ����B, ����B��)���� �����մϴ�."
            strReturn(2) = ""
            
'
        Case "J"        '> ����
            ReDim strReturn(3)
            strReturn(0) = "���ι���(����B, ����A, ����B��)/�ڿ���(����A, ����B, ����B��)���� �����մϴ�."
            strReturn(1) = "���ι��� ��ȸŽ�� �� �����, ��������, ���ƽþƻ�, ��Ȱ�� ���� �� ��2�ܱ���� ���Թݿ��� �����Ͽ� ������ �� �ֽ��ϴ�."
            strReturn(2) = "���ڿ��� ����Ž�� �� ������, �������Х��� ���Թݿ��� �����Ͽ� ������ �� �ֽ��ϴ�."
            strReturn(3) = "�����ð����� ��û�� ���� �ؼҼ��� ��� �������� ���� ���� �ֽ��ϴ�. "
            
        Case "B"        '> �λ�
            ReDim strReturn(2)
            strReturn(0) = "�� �����н��� �ι���� 6���� �� 3������ ������ �� ������, ����, ��������, �����, ������ȸ, ��������, ��2�ܱ��� ������ ���� ���չݺ��� �����մϴ�."
            strReturn(1) = "�� �����н��� �ڿ���� 4���� �� 3������ ������ �� ������, ����II(4����), �������� ���ð��� ����, Ȯ������ ���� ���չݺ��� �����մϴ�."
            strReturn(2) = ""
            
    End Select
    
    Get_StrGongji = strReturn
    
End Function

Public Function Get_StrGongjiJonghab() As String()
    Dim strReturn() As String
    
    '>> �г⺰ ����
    
    Select Case Trim(basModule.SchCD)
    
        Case "N"
            ReDim strReturn(3)
            strReturn(0) = "�� ����, ����, ����, ���� �� (��ȭ)�Ǵ�(�⺻) ���� 1������ �����ؾ� �ϸ�, Ž������ 4���� �� 1������ ������ �� �ֽ��ϴ�."
            strReturn(1) = "�� �ι���� ��Ȱ�� ����, ������ ���, ��������, ���ƽþƻ�, �����, ����, ��2�ܱ���, �ڿ���� ���Х�(4����)�� ��� ���Թݺ��� �����մϴ�."
            strReturn(2) = "�� �ݴ� ������ �� ������ ���� �й� �Ǵ� �չ��� �� �ֽ��ϴ�."
            strReturn(3) = "�� �ι����(����B, ����A, ����B) / �ڿ���(����A, ����B, ����B��)���� �����մϴ�."
           
        Case "K", "W", "Q"
            ReDim strReturn(2)
            strReturn(0) = "���ι��� ��ȸŽ�� �� ���� ��ġ, �����, ��������, ���ƽþƻ�, ������ ���, ��Ȱ�� ���� �� ��2�ܱ���� ���Թݿ��� �����Ͽ� ������ �� �ֽ��ϴ�."
            strReturn(1) = "���ڿ��� ����Ž�� �� ������, ȭ�Х�, ������Х�, �������Х��� ���Թݿ��� �����Ͽ� ������ �� �ֽ��ϴ�."
            strReturn(2) = "���ι���(����B, ����A, ����B��)/�ڿ���(����A, ����B, ����B��)���� �����մϴ�."
            
            
        Case "S"
            ReDim strReturn(3)
            strReturn(0) = "�� ����, ����, ����, ���� �� (��ȭ)�Ǵ�(�⺻) ���� 2������ �����ؾ� �ϸ�, Ž�������� 1������ ������ �� �ֽ��ϴ�."
            strReturn(1) = "�� �ι���� ��Ȱ�� ����, ������ ���, ��������, ���ƽþƻ�, �����, ����, ��2�ܱ���, �ڿ���� ���Х�(4����)�� ��� ���Թݺ��� �����մϴ�."
            strReturn(2) = "�� �ݴ� ������ �� ������ ���� �й� �Ǵ� �չ��� �� �ֽ��ϴ�."
            strReturn(3) = "�� �ι����(����B, ����A, ����B) / �ڿ���(����A, ����B, ����B��)���� �����մϴ�."
                  
        Case "P"
            ReDim strReturn(1)
            strReturn(0) = ""
            strReturn(1) = ""
                             
         Case "M"
            ReDim strReturn(2)
            strReturn(0) = "���ι��� ��ȸŽ�� �� �����, ��������, ���ƽþƻ�, ��Ȱ�� ���� �� ��2�ܱ���� ���Թݿ��� �����Ͽ� ������ �� �ֽ��ϴ�."
            strReturn(1) = "���ι���(����B, ����A, ����B��)/�ڿ���(����A, ����B, ����B��)���� �����մϴ�."
            strReturn(2) = ""
            
'
        Case "J"        '> ����
            ReDim strReturn(3)
            strReturn(0) = "���ι���(����B, ����A, ����B��)/�ڿ���(����A, ����B, ����B��)���� �����մϴ�."
            strReturn(1) = "���ι��� ��ȸŽ�� �� �����, ��������, ���ƽþƻ�, ��Ȱ�� ���� �� ��2�ܱ���� ���Թݿ��� �����Ͽ� ������ �� �ֽ��ϴ�."
            strReturn(2) = "���ڿ��� ����Ž�� �� ������, �������Х��� ���Թݿ��� �����Ͽ� ������ �� �ֽ��ϴ�."
            strReturn(3) = "�����ð����� ��û�� ���� �ؼҼ��� ��� �������� ���� ���� �ֽ��ϴ�. "
            
        Case "B"        '> �λ�
            ReDim strReturn(2)
            strReturn(0) = "�� �����н��� �ι���� 6���� �� 3������ ������ �� ������, ����, ��������, �����, ������ȸ, ��������, ��2�ܱ��� ������ ���� ���չݺ��� �����մϴ�."
            strReturn(1) = "�� �����н��� �ڿ���� 4���� �� 3������ ������ �� ������, ����II(4����), �������� ���ð��� ����, Ȯ������ ���� ���չݺ��� �����մϴ�."
            strReturn(2) = ""
            
    End Select
    
    Get_StrGongjiJonghab = strReturn
    
End Function

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'���� ���� SQL��
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Function Get_SqlKaeyolDecode()
    Dim sStr    As String
    
    sStr = ""
    If Trim(basModule.SchCD) = "N" Then
        sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
        sStr = sStr & "                   '02','�ڿ�',"
        sStr = sStr & "                   '03','��ü',"
        sStr = sStr & "                   '04','����(��)',"
        sStr = sStr & "                   '05','�ι�����',"
        sStr = sStr & "                   '06','�ڿ�����',"
        
        sStr = sStr & "                   '06','�ڿ�����',"
        sStr = sStr & "                   '07','�ż��ι�',"
        sStr = sStr & "                   '08','�ż��ڿ�',"
        sStr = sStr & "                   '09','�ż������ι�',"
        sStr = sStr & "                   '10','�ż������ڿ�',"
        
        sStr = sStr & "                   '11','��)�ι�',"
        sStr = sStr & "                   '12','��)�ڿ�',"
        sStr = sStr & "                   '13','��)��ü',"
        sStr = sStr & "                   '14','��)����(��)',"
        sStr = sStr & "                   '15','��)�ι�����',"
        sStr = sStr & "                   '16','��)�ڿ�����',"
        sStr = sStr & "                   '21','������ι�',"
        sStr = sStr & "                   '22','������ι�'"
        sStr = sStr & "            ) AS GAEYUL,"
        
    '<< �迭 >> : 2008.01.10/ 2008.03.24
    ElseIf Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then
        sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
        sStr = sStr & "                   '02','�ڿ�',"
        
        sStr = sStr & "                   '04','�ָ�����',"
        sStr = sStr & "                   '05','�ָ��Ǵ�',"
        sStr = sStr & "                   '06','�߰�����',"
        sStr = sStr & "                   '07','�߰��Ǵ�',"
        
        sStr = sStr & "                   '11','�������ι�',"
        sStr = sStr & "                   '12','�������ڿ�',"
        
        sStr = sStr & "                   '16','�������ι�16',"
        sStr = sStr & "                   '17','�������ڿ�17',"
        
        sStr = sStr & "                   '19','���ſ�����ι�',"
        sStr = sStr & "                   '20','���ſ�����ڿ�'"
        
        sStr = sStr & "            ) AS GAEYUL,"
        
    '<< �迭 >> : 2008.02.15
    ElseIf Trim(basModule.SchCD) = "S" Then
       sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
        sStr = sStr & "                   '02','�ڿ�',"
        
        sStr = sStr & "                   '03','��ü��',"
        
        sStr = sStr & "                   '05','�����ι�',"
        sStr = sStr & "                   '06','�����ڿ�',"
        
        sStr = sStr & "                   '11','�ż��ι�',"
        sStr = sStr & "                   '12','�ż��ڿ�',"
        
        sStr = sStr & "                   '18','�ι������̾�',"
        sStr = sStr & "                   '19','�ڿ������̾�',"
        sStr = sStr & "                   '21','�����Ư���ι�',"
        sStr = sStr & "                   '22','�����Ư���ڿ�',"
        sStr = sStr & "                   '23','�߰�������ι�',"
        sStr = sStr & "                   '24','�߰�������ڿ�'"
        
        sStr = sStr & "            ) AS GAEYUL,"
    ElseIf Trim(basModule.SchCD) = "J" Then                 '< ����
        sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
        sStr = sStr & "                   '02','�ڿ�',"
        sStr = sStr & "                   '11','�ż��ι�',"
        sStr = sStr & "                   '12','�ż��ڿ�',"
        
        sStr = sStr & "                   '18','�ι������̾�',"
        sStr = sStr & "                   '19','�ڿ������̾�'"
        
        sStr = sStr & "            ) AS GAEYUL,"
        
    ElseIf Trim(basModule.SchCD) = "P" Then                 '< ����
        sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
        sStr = sStr & "                   '02','�ڿ�',"
        sStr = sStr & "                   '03','Ư���ι�',"
        sStr = sStr & "                   '04','Ư���ڿ�'"
        sStr = sStr & "            ) AS GAEYUL,"
        
    ElseIf Trim(basModule.SchCD) = "B" Then                 '< �λ� : 2009.01.09
        sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
        sStr = sStr & "                   '02','�ڿ�',"
        sStr = sStr & "                   '23','�ι�PS',"
        sStr = sStr & "                   '24','�ڿ�PM',"
        sStr = sStr & "                   '05','Ư���ι�',"
        sStr = sStr & "                   '06','Ư���ڿ�',"
        sStr = sStr & "                   '07','������ι�',"
        sStr = sStr & "                   '08','������ڿ�',"
        sStr = sStr & "                   '09','��ȭ�ι�',"
        sStr = sStr & "                   '10','��ȭ�ڿ�'"
        sStr = sStr & "            ) AS GAEYUL,"
        
    Else
        sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
        sStr = sStr & "                   '02','�ڿ�'"
        sStr = sStr & "            ) AS GAEYUL,"
    End If
    
    Get_SqlKaeyolDecode = sStr
End Function

Public Function AddSQL_ClinicToExcel()
    Dim sStr    As String
    
    sStr = ""
    sStr = sStr & " CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '101') > 0 THEN '" & g_sClinic_Ls(0) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '102') > 0 THEN '" & g_sClinic_Ls(1) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '103') > 0 THEN '" & g_sClinic_Ls(2) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '104') > 0 THEN '" & g_sClinic_Ls(3) & "'"
    sStr = sStr & " END END END END ����Ŭ����,"
    sStr = sStr & " CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '111') > 0 THEN '" & g_sClinic_Ms(0) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '112') > 0 THEN '" & g_sClinic_Ms(1) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '113') > 0 THEN '" & g_sClinic_Ms(2) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '114') > 0 THEN '" & g_sClinic_Ms(3) & "'"
    sStr = sStr & " END END END END ����Ŭ����,"
    sStr = sStr & " CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '121') > 0 THEN '" & g_sClinic_Es(0) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '122') > 0 THEN '" & g_sClinic_Es(1) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '123') > 0 THEN '" & g_sClinic_Es(2) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '124') > 0 THEN '" & g_sClinic_Es(3) & "'"
    sStr = sStr & " END END END END ����Ŭ����"
    
    AddSQL_ClinicToExcel = sStr
End Function

'�л� ���� ���� ������ (�뷮��,����)
Public Function Get_StdExcuteSqlToExcel_N(kaeyol As String, Optional day1 As String, Optional day2 As String) As String
    Dim sStr        As String
    
    
    sStr = ""
    sStr = sStr & "  SELECT SCHNO AS �ý����ڵ�   , "
    sStr = sStr & "         ACID  AS �п�   , "
    sStr = sStr & "         EXMID AS �����ȣ, STDNM AS �л�, "
    
    sStr = sStr & " Birth_ymd as �������, "
    
    sStr = sStr & "         DECODE(EXMTYPE,'0','������','1','������') AS ��������, "
    sStr = sStr & "         DECODE(KAEYOL,'01','�ι�',"
    sStr = sStr & "                       '02','�ڿ�',"
'<< �迭 >> : 2008.01.09
    If Trim(basModule.SchCD) = "N" Then
        sStr = sStr & "                   '03','��ü',"
        sStr = sStr & "                   '04','����(��)',"
        sStr = sStr & "                   '05','�ι�����',"
        sStr = sStr & "                   '06','�ڿ�����',"
        
        sStr = sStr & "                   '07','�ż��ι�',"
        sStr = sStr & "                   '08','�ż��ڿ�',"
        sStr = sStr & "                   '09','�ż������ι�',"
        sStr = sStr & "                   '10','�ż������ڿ�',"
        
        sStr = sStr & "                   '11','��)�ι�',"
        sStr = sStr & "                   '12','��)�ڿ�',"
        sStr = sStr & "                   '13','��)��ü',"
        sStr = sStr & "                   '14','��)����(��)',"
        sStr = sStr & "                   '15','��)�ι�����',"
        sStr = sStr & "                   '16','��)�ڿ�����',"
        sStr = sStr & "                   '21','������ι�',"
        sStr = sStr & "                   '22','������ڿ�',"
    End If
'<< �迭 >> : 2008.01.10
    If Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then
        sStr = sStr & "                   '04','�ָ�����',"
        sStr = sStr & "                   '05','�ָ��Ǵ�',"
        sStr = sStr & "                   '06','�߰�����',"
        sStr = sStr & "                   '07','�߰��Ǵ�',"
    
        sStr = sStr & "                   '11','�������ι�',"
        sStr = sStr & "                   '12','�������ڿ�',"
        
        sStr = sStr & "                   '16','�������ι�16',"
        sStr = sStr & "                   '17','�������ڿ�17',"
        
        sStr = sStr & "                   '19','���ſ�����ι�',"
        sStr = sStr & "                   '20','���ſ�����ڿ�',"
        

    End If
'<< �迭 >> : 2008.02.15
    If Trim(basModule.SchCD) = "S" Then
        sStr = sStr & "                   '03','��ü��',"
        'sStr = sStr & "                   '04','Ư���ڿ�',"
        
        sStr = sStr & "                   '05','�����ι�',"
        sStr = sStr & "                   '06','�����ڿ�',"
        
        sStr = sStr & "                   '11','�ż��ι�',"
        sStr = sStr & "                   '12','�ż��ڿ�',"
        
        sStr = sStr & "                   '18','�ι������̾�',"
        sStr = sStr & "                   '19','�ڿ������̾�',"
        sStr = sStr & "                   '21','�����Ư���ι�',"
        sStr = sStr & "                   '22','�����Ư���ڿ�',"
        sStr = sStr & "                   '23','�߰�������ι�',"
        sStr = sStr & "                   '24','�߰�������ڿ�',"
        
    End If
'<< �迭 >> : 2008.02.15
    If Trim(basModule.SchCD) = "P" Then         '< ����
        sStr = sStr & "                   '03','Ư���ι�',"
        sStr = sStr & "                   '04','Ư���ڿ�',"
    End If
    
    If Trim(basModule.SchCD) = "J" Then         '< ����
        sStr = sStr & "                   '11','�ż��ι�',"
        sStr = sStr & "                   '12','�ż��ڿ�',"
        
        sStr = sStr & "                   '18','�ι������̾�',"
        sStr = sStr & "                   '19','�ڿ������̾�',"
    End If
    
'<< �迭 >> : 2009.01.09
    If Trim(basModule.SchCD) = "B" Then         '< �λ�
        sStr = sStr & "                   '05','�����ι�',"
        sStr = sStr & "                   '06','�����ڿ�',"
        sStr = sStr & "                   '07','������ι�',"
        sStr = sStr & "                   '08','������ڿ�',"
        sStr = sStr & "                   '09','��ȭ�ι�',"
        sStr = sStr & "                   '10','��ȭ�ڿ�',"
    End If
    
    sStr = sStr & "                       '','��Ÿ') AS �迭,"
    
    sStr = sStr & "     /* ��Ž, ��Ž �и� */"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(0) & "|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "             '" & constSatams(0) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "             '��1'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS Ž��1,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(1) & "|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "             '" & constSatams(1) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* ��Ž-ȭ��1 */"
    sStr = sStr & "             'ȭ1'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS Ž��2,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(2) & "|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "             '" & constSatams(2) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* ��Ž-�������1 */"
    sStr = sStr & "             '��1'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS Ž��3,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(3) & "|') > 0 THEN          /* ��Ž-�ѱ������� */"
    sStr = sStr & "             '" & constSatams(3) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* ��Ž-��������1 */"
    sStr = sStr & "             '��1'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS Ž��4,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(4) & "|') > 0 THEN          /* ��Ž-����� */"
    sStr = sStr & "             '" & constSatams(4) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "             '��2'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS Ž��5,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(5) & "|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "             '" & constSatams(5) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* ��Ž-ȭ��2 */"
    sStr = sStr & "             'ȭ2'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS Ž��6,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(6) & "|') > 0 THEN          /* ��Ž-�ѱ����� */"
    sStr = sStr & "             '" & constSatams(6) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* ��Ž-�������2 */"
    sStr = sStr & "             '��2'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS Ž��7,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(7) & "|') > 0 THEN          /* ��Ž-��ġ */"
    sStr = sStr & "             '" & constSatams(7) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* ��Ž-��������2 */"
    sStr = sStr & "             '��2'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS Ž��8,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(8) & "|') > 0 THEN          /* ��Ž-��ȸ��ȭ */"
    sStr = sStr & "             '" & constSatams(8) & "'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END AS Ž��9,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(9) & "|') > 0 THEN          /* ��Ž-������ȸ */"
    sStr = sStr & "             '" & constSatams(9) & "'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END AS Ž��10,"
    sStr = sStr & " '' AS Ž��11,"
    
    sStr = sStr & "  "
    sStr = sStr & "      /* ��2�ܱ��� & ���� */"
    sStr = sStr & "              CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '�Ͼ�'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '�Ҿ�'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '�߾�'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '�ѹ�'"
    
    '<< ���� >> : 2008.01.09
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'37|') > 0 THEN '���'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'38|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'39|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'40|') > 0 THEN '�����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'41|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'42|') > 0 THEN '�ƶ���'"
    
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '�̻�'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN 'Ȯ��'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '����'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END END END END END END END END END END END END END END END ��2����,"
    sStr = sStr & "  "
    sStr = sStr & "      /* ��� */"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'91|') > 0 THEN         /* ��� */"
    sStr = sStr & "             '���'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END �����,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'92|') > 0 THEN         /* ���� */"
    sStr = sStr & "             '����'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END �������,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* �ܱ��� */"      '< ����
    sStr = sStr & "             '�ܱ���'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END ��Ž���,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /*  */"            '< ����
    sStr = sStr & "             ' '"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END ��Ž���,"
    sStr = sStr & "  "
    sStr = sStr & "         CY_ACNT AS �������, TOT_AMT AS ��ü�ݾ�    ,"
    sStr = sStr & "         NVL(BASE_AMT1    ,0) AS �⺻�ݾ�1  ,"
    sStr = sStr & "         NVL(BASE_AMT2    ,0) AS �⺻�ݾ�2  ,"
    sStr = sStr & "         NVL(BASE_AMT3    ,0) AS �⺻�ݾ�3  ,"
    sStr = sStr & "         NVL(BASE_AMT4    ,0) AS �⺻�ݾ�4  ,"
    sStr = sStr & "         NVL(BASE_AMT5    ,0) AS �⺻�ݾ�5  ,"
    sStr = sStr & "         NVL(BASE_AMT6    ,0) AS �⺻�ݾ�6  ,"
    sStr = sStr & "         NVL(BASE_AMT7    ,0) AS �⺻�ݾ�7  ,"
    sStr = sStr & "         NVL(BASE_AMT8    ,0) AS �⺻�ݾ�8  ,"
    sStr = sStr & "         NVL(TAMGU_AMT1   ,0) AS Ž�������ݾ�1 ,"
    sStr = sStr & "         NVL(TAMGU_AMT2   ,0) AS Ž�������ݾ�2 ,"
    sStr = sStr & "         NVL(TAMGU_AMT3   ,0) AS Ž�������ݾ�3 ,"
    sStr = sStr & "         NVL(TAMGU_AMT4   ,0) AS Ž�������ݾ�4 ,"
    sStr = sStr & "         NVL(TAMGU_AMT5   ,0) AS Ž�������ݾ�5 ,"
    sStr = sStr & "         NVL(TAMGU_AMT6   ,0) AS Ž�������ݾ�6 ,"
    sStr = sStr & "         NVL(TAMGU_AMT7   ,0) AS Ž�������ݾ�7 ,"
    sStr = sStr & "         NVL(TAMGU_AMT8   ,0) AS Ž�������ݾ�8 ,"
    sStr = sStr & "         NVL(TAMGU_AMT9   ,0) AS Ž�������ݾ�9 ,"
    sStr = sStr & "         NVL(TAMGU_AMT10  ,0) AS Ž�������ݾ�10,"
    sStr = sStr & "         NVL(TAMGU_AMT11  ,0) AS Ž�������ݾ�11,"
    
    sStr = sStr & "      /* Ž�� ���� ������ ó��.. */"
    sStr = sStr & "              CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(0) & "') > 0 THEN '" & constSatams(0) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(1) & "') > 0 THEN '" & constSatams(1) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(2) & "') > 0 THEN '" & constSatams(2) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(3) & "') > 0 THEN '" & constSatams(3) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(4) & "') > 0 THEN '" & constSatams(4) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(5) & "') > 0 THEN '" & constSatams(5) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(6) & "') > 0 THEN '" & constSatams(6) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(7) & "') > 0 THEN '" & constSatams(7) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(8) & "') > 0 THEN '" & constSatams(8) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(9) & "') > 0 THEN '" & constSatams(9) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'51') > 0 THEN '��I'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'52') > 0 THEN 'ȭI'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'53') > 0 THEN '��I'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'54') > 0 THEN '��I'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'55') > 0 THEN '��II'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'56') > 0 THEN 'ȭII'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'57') > 0 THEN '��II'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'58') > 0 THEN '��II'"
    sStr = sStr & "         END END END END END END END END END END END END END END END END END END SEL_X6,"
    
    sStr = sStr & "         K_NUM AS �������, M_NUM AS ��������, E_NUM AS ��������, "
    sStr = sStr & "         (NVL(K_NUM,0)+NVL(M_NUM,0)+NVL(E_NUM,0)) AS ��ü����,"
    sStr = sStr & "         N_NUM AS ���ŵ��,"
    
    
    sStr = sStr & "         DECODE(SEL1_SCH,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS ��1����,"
    sStr = sStr & "         DECODE(SEL2_SCH,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS ��2����,"
    
    sStr = sStr & "         DECODE(PASS1,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS �հ�1   ,"
    sStr = sStr & "         DECODE(PASS2,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS �հ�2   ,"
    sStr = sStr & "         DECODE(PASS3,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS �հ�3   ,"
    sStr = sStr & "         DECODE(PASS4,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS �հ�4   ,"
    
    
    sStr = sStr & "         DECODE(SEX,'M','��','F','��') AS ����        , "
    sStr = sStr & "         SUBSTR(ZIP,1,3)||'-'||SUBSTR(ZIP,4,3) AS �����ȣ, ADDR1 AS �����ּ�      , ADDR2 AS ���ּ�     ,"
    sStr = sStr & "         TEL AS ��ȭ��ȣ, CEL AS �ڵ���        , EMAIL AS �̸���     ,"
    sStr = sStr & "         HIGH_SCH AS ����б� , GRADE_YEAR AS �����⵵ ,"
    sStr = sStr & "         PRNT_NM AS �кθ�� , DECODE(PRNT_RLTN,'1','��','2','��','3','��Ÿ') AS �кθ����, "
    sStr = sStr & "         SUBSTR(PRNT_ZIP,1,3)||'-'||SUBSTR(PRNT_ZIP,4,3) AS �кθ�_�����ȣ, PRNT_ADDR1 AS �кθ�_�����ּ� , PRNT_ADDR2 AS �кθ�_���ּ�,"
    sStr = sStr & "         PRNT_TEL AS �кθ�_��ȭ��ȣ  , PRNT_CEL AS �кθ�_�ڵ���   , PRNT_JOB AS �кθ�_����   , PRNT_W_TEL AS �кθ�_������ȭ ,"
    sStr = sStr & "         PHOTO_PATH AS �����������, "
    sStr = sStr & "         DECODE(R_WAY,'1','�п����','2','���ͳݵ��','3','�п����') AS ��Ϲ�ȣ, "
    sStr = sStr & "         ORD_NO AS �ֹ���ȣ, "
    sStr = sStr & "         ACID||EXMID AS �̹������ϸ�, "
    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','',ACID) AS WANT_ACID "
    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','" & Trim(basModule.SchCD) & "',ACID) AS WANT_ACID, "       '< TEST
    sStr = sStr & "         REGDATE AS �������, GET_PAYGUBN(ORD_NO) AS ������, CASH_BILL_NUM AS ���ݿ�����,"
    sStr = sStr & "         DECODE(MU_TYPE,'1','������','2','6�� �򰡿�','3','9�� �򰡿�','4','6�� �򰡿�','5','9','���ŵ��','9�� �򰡿�','') AS ���, "
    sStr = sStr & "         CL_CLOSE AS �Ϸ��� "
    
    sStr = sStr & " , "
        sStr = sStr & "        J01 AS ���          ,"
        sStr = sStr & "        K01 AS ���_��       ,"
        sStr = sStr & "        J02 AS ������        ,"
        sStr = sStr & "        K02 AS ��������_��   ,"
        sStr = sStr & "        J03 AS �ܱ���        ,"
        sStr = sStr & "        K03 AS �ܱ���_��     ,"
                                   
        sStr = sStr & "        J04 AS " & constSatams(0) & "_��1      ,"
        sStr = sStr & "        K04 AS " & constSatams(0) & "_��1_��   ,"
        sStr = sStr & "        J05 AS " & constSatams(1) & "_ȭ1      ,"
        sStr = sStr & "        K05 AS " & constSatams(1) & "_ȭ1_��   ,"
        sStr = sStr & "        J06 AS " & constSatams(2) & "_��1      ,"
        sStr = sStr & "        K06 AS " & constSatams(2) & "_��1_��   ,"
        sStr = sStr & "        J07 AS " & constSatams(3) & "_����1    ,"
        sStr = sStr & "        K07 AS " & constSatams(3) & "_����1_�� ,"
        sStr = sStr & "        J08 AS " & constSatams(4) & "_��2      ,"
        sStr = sStr & "        K08 AS " & constSatams(4) & "_��2_��   ,"
        sStr = sStr & "        J09 AS " & constSatams(5) & "_ȭ2      ,"
        sStr = sStr & "        K09 AS " & constSatams(5) & "_ȭ2_��   ,"
        sStr = sStr & "        J10 AS " & constSatams(6) & "_��2      ,"
        sStr = sStr & "        K10 AS " & constSatams(6) & "_��2_��   ,"
        sStr = sStr & "        J11 AS " & constSatams(7) & "_����2    ,"
        sStr = sStr & "        K11 AS " & constSatams(7) & "_����2_�� ,"
                                   
        sStr = sStr & "        J12 AS " & constSatams(8) & "          ,"
        sStr = sStr & "        K12 AS " & constSatams(8) & "_��       ,"
        sStr = sStr & "        J13 AS " & constSatams(9) & "          ,"
        sStr = sStr & "        K13 AS " & constSatams(9) & "_��       ,"
        sStr = sStr & " ' ' AS K14, "
        sStr = sStr & " ' ' AS J14, "
        sStr = sStr & "        J15 AS ����_����     ,"
        sStr = sStr & "        K15 AS ����_����_��  ,"
        sStr = sStr & "        J16 AS �Ͼ�_�̻�     ,"
        sStr = sStr & "        K16 AS �Ͼ�_�̻�_��  ,"
        sStr = sStr & "        J17 AS ����_Ȯ��     ,"
        sStr = sStr & "        K17 AS ����_Ȯ��_��  ,"
        sStr = sStr & "        J18 AS �Ҿ�_������   ,"
        sStr = sStr & "        K18 AS �Ҿ�_������_��,"
                                   
        sStr = sStr & "        J19 AS �߾�          ,"
        sStr = sStr & "        K19 AS �߾�_��       ,"
        sStr = sStr & "        J20 AS �ѹ�          ,"
        sStr = sStr & "        K20 AS �ѹ�_��       ,"
        sStr = sStr & "        J21 AS �ƶ���        ,"
        sStr = sStr & "        K21 AS �ƶ���_��     ,"
        
        ' �뷮�� ��û�� ���� �����ܴ�... �׷��� �ѿ����� ������ �߰��س����̴�.. �׷��� ������ �ȵǾ�����.
        ' �ؿ��ٰ���.. ����..
        sStr = sStr & "        D_UNIVCD AS ��������, D_MAJORCD AS �����ܴ�, "
        
        ' Ŭ���� �߰�
        sStr = sStr & " CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '101') > 0 THEN '" & g_sClinic_Ls(0) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '102') > 0 THEN '" & g_sClinic_Ls(1) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '103') > 0 THEN '" & g_sClinic_Ls(2) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '104') > 0 THEN '" & g_sClinic_Ls(3) & "'"
        sStr = sStr & " END END END END ����Ŭ����,"
        sStr = sStr & " CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '111') > 0 THEN '" & g_sClinic_Ms(0) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '112') > 0 THEN '" & g_sClinic_Ms(1) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '113') > 0 THEN '" & g_sClinic_Ms(2) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '114') > 0 THEN '" & g_sClinic_Ms(3) & "'"
        sStr = sStr & " END END END END ����Ŭ����,"
        sStr = sStr & " CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '121') > 0 THEN '" & g_sClinic_Es(0) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '122') > 0 THEN '" & g_sClinic_Es(1) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '123') > 0 THEN '" & g_sClinic_Es(2) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '124') > 0 THEN '" & g_sClinic_Es(3) & "'"
        sStr = sStr & " END END END END ����Ŭ����"
        
        sStr = sStr & "    FROM ( "
    
            sStr = sStr & "  SELECT A.SCHNO           ,"
            sStr = sStr & "         MAX(ACID      ) AS ACID       ,"
            sStr = sStr & "         MAX(EXMID     ) AS EXMID      ,"
            sStr = sStr & "         MAX(STDNM     ) AS STDNM      ,"
            sStr = sStr & "         MAX(D_UNIVCD     ) AS D_UNIVCD      ,"
            sStr = sStr & "         MAX(D_MAJORCD     ) AS D_MAJORCD      ,"
            sStr = sStr & "         MAX(Birth_ymd     ) AS Birth_ymd      ,"
            sStr = sStr & "         MAX(EXMTYPE   ) AS EXMTYPE    , MAX(KAEYOL    ) AS KAEYOL     ,"
            sStr = sStr & "         MAX(SEL1      ) AS SEL1       , MAX(SEL2      ) AS SEL2       , MAX(SEL3      ) AS SEL3      , MAX(SEL4      ) AS SEL4      , MAX(SEL5      ) AS  SEL5 , MAX(SEL6      ) AS  SEL6, MAX(SEL7      ) AS  SEL7           ,"
            sStr = sStr & "         MAX(K_NUM     ) AS K_NUM      , MAX(M_NUM     ) AS M_NUM      , MAX(E_NUM     ) AS E_NUM     , MAX(N_NUM     ) AS N_NUM     , MAX(TOT_NUM   ) AS TOT_NUM   ,"
            sStr = sStr & "         MAX(SEL1_SCH  ) AS SEL1_SCH   , MAX(SEL2_SCH  ) AS SEL2_SCH   ,"
            sStr = sStr & "         MAX(PASS1     ) AS PASS1      , MAX(PASS2     ) AS PASS2      , MAX(PASS3     ) AS PASS3     , MAX(PASS4     ) AS PASS4     , MAX(CL_CLOSE  ) AS  CL_CLOSE  ,"
            sStr = sStr & "         MAX(CY_ACNT   ) AS CY_ACNT    , MAX(TOT_AMT   ) AS TOT_AMT    ,"
            sStr = sStr & "         MAX(BASE_AMT1 ) AS BASE_AMT1  , MAX(BASE_AMT2 ) AS BASE_AMT2  , MAX(BASE_AMT3 ) AS BASE_AMT3 , MAX(BASE_AMT4 ) AS BASE_AMT4 ,"
            sStr = sStr & "         MAX(BASE_AMT5 ) AS BASE_AMT5  , MAX(BASE_AMT6 ) AS BASE_AMT6  , MAX(BASE_AMT7 ) AS BASE_AMT7 , MAX(BASE_AMT8 ) AS BASE_AMT8 ,"
            sStr = sStr & "         MAX(TAMGU_AMT1) AS TAMGU_AMT1 , MAX(TAMGU_AMT2) AS TAMGU_AMT2 , MAX(TAMGU_AMT3) AS TAMGU_AMT3, MAX(TAMGU_AMT4) AS TAMGU_AMT4, MAX(TAMGU_AMT5) AS  TAMGU_AMT5,"
            sStr = sStr & "         MAX(TAMGU_AMT6) AS TAMGU_AMT6 , MAX(TAMGU_AMT7) AS TAMGU_AMT7 , MAX(TAMGU_AMT8) AS TAMGU_AMT8, MAX(TAMGU_AMT9) AS TAMGU_AMT9, MAX(TAMGU_AMT10) AS TAMGU_AMT10, MAX(TAMGU_AMT11) AS TAMGU_AMT11,"
            sStr = sStr & "         MAX(SEX       ) AS SEX        ,"
            sStr = sStr & "         MAX(ZIP       ) AS ZIP        , MAX(ADDR1     ) AS ADDR1      , MAX(ADDR2     ) AS ADDR2     ,"
            sStr = sStr & "         MAX(TEL       ) AS TEL        , MAX(CEL       ) AS CEL        , MAX(EMAIL     ) AS EMAIL     ,"
            sStr = sStr & "         MAX(HIGH_SCH  ) AS HIGH_SCH   , MAX(GRADE_YEAR) AS GRADE_YEAR ,"
            sStr = sStr & "         MAX(PRNT_NM   ) AS PRNT_NM    , MAX(PRNT_RLTN ) AS PRNT_RLTN  ,"
            sStr = sStr & "         MAX(PRNT_ZIP  ) AS PRNT_ZIP   , MAX(PRNT_ADDR1) AS PRNT_ADDR1 , MAX(PRNT_ADDR2) AS PRNT_ADDR2,"
            sStr = sStr & "         MAX(PRNT_TEL  ) AS PRNT_TEL   , MAX(PRNT_CEL  ) AS PRNT_CEL   , MAX(PRNT_JOB  ) AS PRNT_JOB  , MAX(PRNT_W_TEL) AS PRNT_W_TEL,"
            sStr = sStr & "         MAX(PHOTO_PATH) AS PHOTO_PATH , MAX(R_WAY     ) AS R_WAY      , MAX(ORD_NO    ) AS ORD_NO    , "
            sStr = sStr & "         MAX(TO_CHAR(REGDATE,'YYYY-MM-DD HH24:MI:SS')) AS REGDATE      , MAX(MU_TYPE   ) AS MU_TYPE   , MAX(CASH_BILL_NUM) AS CASH_BILL_NUM"
            
                    sStr = sStr & " , "
                    sStr = sStr & "        SUM(J01) AS J01,"
                    sStr = sStr & "        SUM(K01) AS K01,"
                    sStr = sStr & "        SUM(J02) AS J02,"
                    sStr = sStr & "        SUM(K02) AS K02,"
                    sStr = sStr & "        SUM(J03) AS J03,"
                    sStr = sStr & "        SUM(K03) AS K03,"
                    
                    sStr = sStr & "        SUM(J04) AS J04,"
                    sStr = sStr & "        SUM(K04) AS K04,"
                    sStr = sStr & "        SUM(J05) AS J05,"
                    sStr = sStr & "        SUM(K05) AS K05,"
                    sStr = sStr & "        SUM(J06) AS J06,"
                    sStr = sStr & "        SUM(K06) AS K06,"
                    sStr = sStr & "        SUM(J07) AS J07,"
                    sStr = sStr & "        SUM(K07) AS K07,"
                    sStr = sStr & "        SUM(J08) AS J08,"
                    sStr = sStr & "        SUM(K08) AS K08,"
                    sStr = sStr & "        SUM(J09) AS J09,"
                    sStr = sStr & "        SUM(K09) AS K09,"
                    sStr = sStr & "        SUM(J10) AS J10,"
                    sStr = sStr & "        SUM(K10) AS K10,"
                    sStr = sStr & "        SUM(J11) AS J11,"
                    sStr = sStr & "        SUM(K11) AS K11,"
                    
                    sStr = sStr & "        SUM(J12) AS J12,"
                    sStr = sStr & "        SUM(K12) AS K12,"
                    sStr = sStr & "        SUM(J13) AS J13,"
                    sStr = sStr & "        SUM(K13) AS K13,"
                    sStr = sStr & "   ' ' AS J14, "
                    sStr = sStr & "   ' ' AS K14, "
                    
                    sStr = sStr & "        SUM(J15) AS J15,"
                    sStr = sStr & "        SUM(K15) AS K15,"
                    sStr = sStr & "        SUM(J16) AS J16,"
                    sStr = sStr & "        SUM(K16) AS K16,"
                    sStr = sStr & "        SUM(J17) AS J17,"
                    sStr = sStr & "        SUM(K17) AS K17,"
                    sStr = sStr & "        SUM(J18) AS J18,"
                    sStr = sStr & "        SUM(K18) AS K18,"
                    
                    sStr = sStr & "        SUM(J19) AS J19,"
                    sStr = sStr & "        SUM(K19) AS K19,"
                    sStr = sStr & "        SUM(J20) AS J20,"
                    sStr = sStr & "        SUM(K20) AS K20,"
                    sStr = sStr & "        SUM(J21) AS J21,"
                    sStr = sStr & "        SUM(K21) AS K21"
            
            sStr = sStr & "    FROM ("
            '---------------------------------------------------------------------------- ��ü�л� ��ȸ START
            sStr = sStr & "          SELECT *"
            sStr = sStr & "            FROM CLSTD01TB"
            sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "             AND EXMID > ' ' "
            
    If Trim(Right(kaeyol, 30)) <> "ALL" Then
            sStr = sStr & "             AND KAEYOL = '" & Trim(Right(kaeyol, 30)) & "'"
    End If
    
    '<< �Ⱓ���� >>
    If day1 <> "" Or day2 <> "" Then
        If Trim(day1) <> "" And Trim(day2) <> "" Then
            sStr = sStr & "                 AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                                 AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) = "" And Trim(day2) <> "" Then
            sStr = sStr & "                 AND REGDATE BETWEEN TO_DATE('19000101000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                                 AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) <> "" And Trim(day2) = "" Then
            sStr = sStr & "                 AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                                 AND TO_DATE('99991231235959', 'YYYYMMDDHH24MISS') "
        End If
        
    End If
    
            sStr = sStr & "             AND BIGO2 IS NULL "
            sStr = sStr & "          UNION ALL"
            '---------------------------------------------------------------------------- ��ü�л� ��ȸ END
            '---------------------------------------------------------------------------- �հ��� ��ȸ START
            sStr = sStr & "          SELECT *"
            sStr = sStr & "            From CLSTD01TB"
            sStr = sStr & "           WHERE (PASS1 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS2 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS3 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS4 = '" & Trim(basModule.SchCD) & "'" & " )"
            sStr = sStr & "             AND EXMID > ' ' "
    If Trim(Right(kaeyol, 30)) <> "ALL" Then
            sStr = sStr & "             AND KAEYOL = '" & Trim(Right(kaeyol, 30)) & "'"
    End If
    
    
    '<< �Ⱓ���� >>
    If day1 <> "" Or day2 <> "" Then
        If Trim(day1) <> "" And Trim(day2) <> "" Then
            sStr = sStr & "                 AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                                 AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) = "" And Trim(day2) <> "" Then
            sStr = sStr & "                 AND REGDATE BETWEEN TO_DATE('19000101000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                                 AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) <> "" And Trim(day2) = "" Then
            sStr = sStr & "                 AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                                 AND TO_DATE('99991231235959', 'YYYYMMDDHH24MISS') "
        End If
        
    End If
            sStr = sStr & "             AND BIGO2 IS NULL "
            
            sStr = sStr & "          ) A, "
            
            sStr = sStr & "               ("
        
                sStr = sStr & "         SELECT SCHNO,"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J01,    /* ���                  */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K01,    /* �����  ���          */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J02,    /* ��������              */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K02,    /* �����  ��������      */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J03,    /* �ܱ���                */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K03,    /* �����  �ܱ���        */"
                
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(0) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '51', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J04,    /* ��Ž-" & constSatams(0) & "       , ��Ž-����1             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(0) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '51', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K04,    /* �����  ��Ž-" & constSatams(0) & "        , ��Ž-����1     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(1) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '52', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J05,    /* ��Ž-" & constSatams(1) & "        , ��Ž-ȭ��1             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(1) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '52', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K05,    /* �����  ��Ž-" & constSatams(1) & "        , ��Ž-ȭ��1     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '53', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J06,    /* ��Ž-" & constSatams(2) & "        , ��Ž-�������1             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '53', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K06,    /* �����  ��Ž-" & constSatams(2) & "        , ��Ž-�������1     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '54', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J07,    /* ��Ž-" & constSatams(3) & "  , ��Ž-��������1         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '54', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K07,    /* �����  ��Ž-" & constSatams(3) & "  , ��Ž-��������1 */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '55', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J08,    /* ��Ž-" & constSatams(4) & "      , ��Ž-����2             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '55', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K08,    /* �����  ��Ž-" & constSatams(4) & "      , ��Ž-����2     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '56', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J09,    /* ��Ž-" & constSatams(5) & "    , ��Ž-ȭ��2             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '56', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K09,    /* �����  ��Ž-" & constSatams(5) & "    , ��Ž-ȭ��2     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '57', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J10,      /* ��Ž-" & constSatams(6) & "    , ��Ž-�������2           */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '57', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K10,      /* ����� ��Ž-" & constSatams(6) & "    , ��Ž-�������2    */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(7) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '58', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J11,    /* ��Ž-" & constSatams(7) & "        , ��Ž-��������2         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(7) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '58', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K11,    /* �����  ��Ž-" & constSatams(7) & "        , ��Ž-��������2 */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(8) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J12,    /* ��Ž-" & constSatams(8) & "         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(8) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K12,    /* �����  ��Ž-" & constSatams(8) & " */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(9) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J13,    /* ��Ž-" & constSatams(9) & "         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(9) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K13,    /* �����  ��Ž-" & constSatams(9) & " */"
                sStr = sStr & " '' AS K14, "
                sStr = sStr & " '' AS J14, "
                
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_NUM,'X',0, SUB_NUM), '81', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J15,    /* ����             , ������                 */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_BAK,'X',0, SUB_BAK), '81', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K15,    /* �����  ����             , ������         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_NUM,'X',0, SUB_NUM), '82', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J16,    /* �Ͼ�             , �̻����               */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_BAK,'X',0, SUB_BAK), '82', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K16,    /* �����  �Ͼ�             , �̻����       */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_NUM,'X',0, SUB_NUM), '83', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J17,    /* �����ĳ�         , Ȯ�����               */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_BAK,'X',0, SUB_BAK), '83', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K17,    /* �����  �����ĳ�         , Ȯ�����       */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_NUM,'X',0, SUB_NUM), '43', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J18,    /* �Ҿ�             , ��������               */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_BAK,'X',0, SUB_BAK), '43', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K18,    /* �����  �Ҿ�             , ��������       */"
                
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J19,    /* �߱���                */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K19,    /* �����  �߱���        */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J20,    /* �ѹ�                  */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K20,    /* �����  �ѹ�          */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J21,    /* �ƶ���                */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K21     /* �����  �ƶ���        */"
                sStr = sStr & "           FROM CLSTD03TB"
        
        sStr = sStr & "                ) B"
        sStr = sStr & "        WHERE A.SCHNO = B.SCHNO(+)"
            
            sStr = sStr & "   GROUP BY A.SCHNO"
            '---------------------------------------------------------------------------- �հ��� ��ȸ END
    
    sStr = sStr & "    ) "
    sStr = sStr & " ORDER BY EXMID "
    
    Get_StdExcuteSqlToExcel_N = sStr
End Function


'�л� ���� ���� ������ (�뷮��,���� �̿ܿ�)
Public Function Get_StdExcuteSqlToExcel(kaeyol As String, Optional day1 As String, Optional day2 As String) As String
    
    Dim sStr         As String
    Dim ni           As Long
    ni = 0
    
    sStr = ""
    sStr = sStr & "  SELECT  "
    sStr = sStr & "         ACID  AS �п�   , "
    sStr = sStr & "         EXMID AS �����ȣ, STDNM AS �л�,"
    sStr = sStr & "         birth_ymd AS �������, "
    sStr = sStr & "         DECODE(EXMTYPE,'0','������','1','������') AS ��������, "
    sStr = sStr & "         DECODE(KAEYOL,'01','�ι�',"
    sStr = sStr & "                       '02','�ڿ�',"
'<< �迭 >> : 2008.01.09
    If Trim(basModule.SchCD) = "N" Then
        sStr = sStr & "                   '03','��ü',"
        sStr = sStr & "                   '04','����(��)',"
        sStr = sStr & "                   '05','�ι�����',"
        sStr = sStr & "                   '06','�ڿ�����',"
        
        sStr = sStr & "                   '07','�ż��ι�',"
        sStr = sStr & "                   '08','�ż��ڿ�',"
        sStr = sStr & "                   '09','�ż������ι�',"
        sStr = sStr & "                   '10','�ż������ڿ�',"
        
        sStr = sStr & "                   '11','��)�ι�',"
        sStr = sStr & "                   '12','��)�ڿ�',"
        sStr = sStr & "                   '13','��)��ü',"
        sStr = sStr & "                   '14','��)����(��)',"
        sStr = sStr & "                   '15','��)�ι�����',"
        sStr = sStr & "                   '16','��)�ڿ�����',"
        sStr = sStr & "                   '21','������ι�',"
        sStr = sStr & "                   '22','������ڿ�',"
    End If
'<< �迭 >> : 2008.01.10
    If Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then
        sStr = sStr & "                   '04','�ָ�����',"
        sStr = sStr & "                   '05','�ָ��Ǵ�',"
        sStr = sStr & "                   '06','�߰�����',"
        sStr = sStr & "                   '07','�߰��Ǵ�',"
    
        sStr = sStr & "                   '11','�������ι�',"
        sStr = sStr & "                   '12','�������ڿ�',"
        
        sStr = sStr & "                   '16','�������ι�16',"
        sStr = sStr & "                   '17','�������ڿ�17',"
        
        sStr = sStr & "                   '19','���ſ�����ι�',"
        sStr = sStr & "                   '20','���ſ�����ڿ�',"
    End If
'<< �迭 >> : 2008.02.15
    If Trim(basModule.SchCD) = "S" Then
        sStr = sStr & "                   '03','��ü��',"
        'sStr = sStr & "                   '04','Ư���ڿ�',"
        
        sStr = sStr & "                   '05','�����ι�',"
        sStr = sStr & "                   '06','�����ڿ�',"
        
        sStr = sStr & "                   '11','�ż��ι�',"
        sStr = sStr & "                   '12','�ż��ڿ�',"
        
        sStr = sStr & "                   '18','�ι������̾�',"
        sStr = sStr & "                   '19','�ڿ������̾�',"
        
    End If
'<< �迭 >> : 2008.02.15
    If Trim(basModule.SchCD) = "P" Then         '< ����
        sStr = sStr & "                   '03','Ư���ι�',"
        sStr = sStr & "                   '04','Ư���ڿ�',"
    End If
    
    If Trim(basModule.SchCD) = "J" Then         '< ����
        sStr = sStr & "                   '11','�ż��ι�',"
        sStr = sStr & "                   '12','�ż��ڿ�',"
        
        sStr = sStr & "                   '18','�ι������̾�',"
        sStr = sStr & "                   '19','�ڿ������̾�',"
    End If
    
'<< �迭 >> : 2009.01.09
    If Trim(basModule.SchCD) = "B" Then         '< �λ�7
        sStr = sStr & "                   '05','�����ι�',"
        sStr = sStr & "                   '06','�����ڿ�',"
        sStr = sStr & "                   '07','������ι�',"
        sStr = sStr & "                   '08','������ڿ�',"
        sStr = sStr & "                   '09','��ȭ�ι�',"
        sStr = sStr & "                   '10','��ȭ�ڿ�',"
    End If
    
    sStr = sStr & "                       '','��Ÿ') AS �迭,"
    
    sStr = sStr & "     /* ��Ž, ��Ž �и� */"
    For ni = 0 To SATAM_COUNT - 1
        sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(ni) & "|') > 0 THEN          /* ��Ž-" & constSatams(ni) & " */"
        sStr = sStr & "             '" & constSatams(ni) & "'"
        sStr = sStr & "         ELSE "

        If ni < GWATAM_COUNT - 1 Then
            sStr = sStr & "         CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'" & constGwatamCodes(ni) & "|') > 0 THEN     /* ��Ž-" & constGwatams(ni) & " */"
            sStr = sStr & "             '" & constGwatams(ni) & "'"
            sStr = sStr & "         ELSE"
            sStr = sStr & "             ' '"
            sStr = sStr & "         END "
        Else
            sStr = sStr & "         ' '"
        End If
        sStr = sStr & "         END AS Ž��" & CStr(ni) & ", "

    Next ni
    
    
    If basModule.SchCD = "J" Then
        sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & TGANG_CODE & "|') > 0 THEN          /* ��Ž-Ư�� */"
        sStr = sStr & "             'Ư��'"
        sStr = sStr & "         ELSE "
        sStr = sStr & "             CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'" & TGANG_CODE & "|') > 0 THEN     /* ��Ž-Ư��*/"
        sStr = sStr & "                'Ư��'"
        sStr = sStr & "             ELSE"
        sStr = sStr & "                 ' '"
        sStr = sStr & "             END "
        sStr = sStr & "         END AS Ž��11, "
    End If
    

    sStr = sStr & "  "
    sStr = sStr & "      /* ��2�ܱ��� & ���� */"
    sStr = sStr & "              CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '�Ͼ�'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '�Ҿ�'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '�߾�'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '�ѹ�'"
    
    '<< ���� >> : 2008.01.09
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'37|') > 0 THEN '���'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'38|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'39|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'40|') > 0 THEN '�����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'41|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'42|') > 0 THEN '�ƶ���'"
    
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '�̻�'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN 'Ȯ��'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '����'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END END END END END END END END END END END END END END END ��2����,"
    sStr = sStr & "  "
    sStr = sStr & "      /* ��� */"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'91|') > 0 THEN         /* ��� */"
    sStr = sStr & "             '���'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END �����,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'92|') > 0 THEN         /* ���� */"
    sStr = sStr & "             '����'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END �������,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* �ܱ��� */"      '< ����
    sStr = sStr & "             '�ܱ���'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END ��Ž���,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /*  */"            '< ����
    sStr = sStr & "             ' '"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END ��Ž���,"
    sStr = sStr & "  "
    sStr = sStr & "         CY_ACNT AS �������, TOT_AMT AS ��ü�ݾ�    ,"
    sStr = sStr & "         NVL(BASE_AMT1    ,0) AS �⺻�ݾ�1  ,"
    sStr = sStr & "         NVL(BASE_AMT2    ,0) AS �⺻�ݾ�2  ,"
    sStr = sStr & "         NVL(BASE_AMT3    ,0) AS �⺻�ݾ�3  ,"
    sStr = sStr & "         NVL(BASE_AMT4    ,0) AS �⺻�ݾ�4  ,"
    sStr = sStr & "         NVL(BASE_AMT5    ,0) AS �⺻�ݾ�5  ,"
    sStr = sStr & "         NVL(BASE_AMT6    ,0) AS �⺻�ݾ�6  ,"
    sStr = sStr & "         NVL(BASE_AMT7    ,0) AS �⺻�ݾ�7  ,"
    sStr = sStr & "         NVL(BASE_AMT8    ,0) AS �⺻�ݾ�8  ,"
    sStr = sStr & "         NVL(TAMGU_AMT1   ,0) AS Ž�������ݾ�1 ,"
    sStr = sStr & "         NVL(TAMGU_AMT2   ,0) AS Ž�������ݾ�2 ,"
    sStr = sStr & "         NVL(TAMGU_AMT3   ,0) AS Ž�������ݾ�3 ,"
    sStr = sStr & "         NVL(TAMGU_AMT4   ,0) AS Ž�������ݾ�4 ,"
    sStr = sStr & "         NVL(TAMGU_AMT5   ,0) AS Ž�������ݾ�5 ,"
    sStr = sStr & "         NVL(TAMGU_AMT6   ,0) AS Ž�������ݾ�6 ,"
    sStr = sStr & "         NVL(TAMGU_AMT7   ,0) AS Ž�������ݾ�7 ,"
    sStr = sStr & "         NVL(TAMGU_AMT8   ,0) AS Ž�������ݾ�8 ,"
    sStr = sStr & "         NVL(TAMGU_AMT9   ,0) AS Ž�������ݾ�9 ,"
    sStr = sStr & "         NVL(TAMGU_AMT10  ,0) AS Ž�������ݾ�10,"
    sStr = sStr & "         NVL(TAMGU_AMT11  ,0) AS Ž�������ݾ�11,"
    
    sStr = sStr & "         K_NUM AS �������, M_NUM AS ��������, E_NUM AS ��������, "
    sStr = sStr & "         (NVL(K_NUM,0)+NVL(M_NUM,0)+NVL(E_NUM,0)) AS ��ü����, N_NUM AS ���ŵ��, "
    
    
    sStr = sStr & "         DECODE(SEL1_SCH,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS ��1����,"
    sStr = sStr & "         DECODE(SEL2_SCH,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS ��2����,"
    
    sStr = sStr & "         DECODE(PASS1,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS �հ�1   ,"
    sStr = sStr & "         DECODE(PASS2,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS �հ�2   ,"
    sStr = sStr & "         DECODE(PASS3,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS �հ�3   ,"
    sStr = sStr & "         DECODE(PASS4,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS �հ�4   ,"
    
    
    sStr = sStr & "         DECODE(SEX,'M','��','F','��') AS ����        , "
    sStr = sStr & "         SUBSTR(ZIP,1,3)||'-'||SUBSTR(ZIP,4,3) AS �����ȣ, ADDR1 AS �����ּ�      , ADDR2 AS ���ּ�     ,"
    sStr = sStr & "         TEL AS ��ȭ��ȣ, CEL AS �ڵ���        , EMAIL AS �̸���     ,"
    sStr = sStr & "         HIGH_SCH AS ����б� , GRADE_YEAR AS �����⵵ ,"
    sStr = sStr & "         PRNT_NM AS �кθ�� , DECODE(PRNT_RLTN,'1','��','2','��','3','��Ÿ') AS �кθ����, "
    sStr = sStr & "         SUBSTR(PRNT_ZIP,1,3)||'-'||SUBSTR(PRNT_ZIP,4,3) AS �кθ�_�����ȣ, PRNT_ADDR1 AS �кθ�_�����ּ� , PRNT_ADDR2 AS �кθ�_���ּ�,"
    sStr = sStr & "         PRNT_TEL AS �кθ�_��ȭ��ȣ  , PRNT_CEL AS �кθ�_�ڵ���   , PRNT_JOB AS �кθ�_����   , PRNT_W_TEL AS �кθ�_������ȭ ,"
    sStr = sStr & "         PHOTO_PATH AS �����������, "
    sStr = sStr & "         DECODE(R_WAY,'1','�п����','2','���ͳݵ��','3','�п����') AS ��Ϲ�ȣ, "
    sStr = sStr & "         ORD_NO AS �ֹ���ȣ, "
    sStr = sStr & "         ACID||EXMID AS �̹������ϸ�, "
    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','',ACID) AS WANT_ACID "
    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','" & Trim(basModule.SchCD) & "',ACID) AS WANT_ACID, "       '< TEST
    sStr = sStr & "         REGDATE AS �������, GET_PAYGUBN(ORD_NO) AS ������, CASH_BILL_NUM AS ���ݿ�����,"
    sStr = sStr & "         DECODE(MU_TYPE,'1','������','2','6�� �򰡿�','3','9�� �򰡿�','4','6�� �򰡿�','9','���ŵ��','5','9�� �򰡿�','') AS ���, "
    sStr = sStr & "         CL_CLOSE AS �Ϸ��� ,"
    
    Select Case Trim(basModule.SchCD)
        Case "S"
            'sStr = sStr & " DECODE(PTS_SEL,'1','����','2','6�� �򰡿�','3','9�� �򰡿�','4','6�� �򰡿�','5','9�� �򰡿�','') AS ����, "
            sStr = sStr & " DECODE(PTS_SEL,'1','����','2','����','') AS ����, "
        Case "P"
            sStr = sStr & " DECODE(PTS_SEL,'8','����','9','2010 ��','6','3���','','') AS ����, "
        Case Else
            sStr = sStr & " '' AS ����,"
    End Select
    
        sStr = sStr & "        J01 AS ���          ,"
        sStr = sStr & "        K01 AS ���_��       ,"
        sStr = sStr & "        J02 AS ������        ,"
        sStr = sStr & "        K02 AS ��������_��   ,"
        sStr = sStr & "        J03 AS �ܱ���        ,"
        sStr = sStr & "        K03 AS �ܱ���_��     ,"
                                   
        sStr = sStr & "        J04 AS " & constSatams(0) & "_" & constGwatams(0) & "      ,"
        sStr = sStr & "        K04 AS " & constSatams(0) & "_" & constGwatams(0) & "_��   ,"
        sStr = sStr & "        J05 AS " & constSatams(1) & "_" & constGwatams(1) & "      ,"
        sStr = sStr & "        K05 AS " & constSatams(1) & "_" & constGwatams(1) & "_��   ,"
        sStr = sStr & "        J06 AS " & constSatams(2) & "_" & constGwatams(2) & "      ,"
        sStr = sStr & "        K06 AS " & constSatams(2) & "_" & constGwatams(2) & "_��   ,"
        sStr = sStr & "        J07 AS " & constSatams(3) & "_" & constGwatams(3) & "      ,"
        sStr = sStr & "        K07 AS " & constSatams(3) & "_" & constGwatams(3) & "_��   ,"
        sStr = sStr & "        J08 AS " & constSatams(4) & "_" & constGwatams(4) & "      ,"
        sStr = sStr & "        K08 AS " & constSatams(4) & "_" & constGwatams(4) & "_��   ,"
        sStr = sStr & "        J09 AS " & constSatams(5) & "_" & constGwatams(5) & "      ,"
        sStr = sStr & "        K09 AS " & constSatams(5) & "_" & constGwatams(5) & "_��   ,"
        sStr = sStr & "        J10 AS " & constSatams(6) & "_" & constGwatams(6) & "      ,"
        sStr = sStr & "        K10 AS " & constSatams(6) & "_" & constGwatams(6) & "_��   ,"
        sStr = sStr & "        J11 AS " & constSatams(7) & "_" & constGwatams(7) & "      ,"
        sStr = sStr & "        K11 AS " & constSatams(7) & "_" & constGwatams(7) & "_��   ,"
                                   
        sStr = sStr & "        J12 AS " & constSatams(8) & "          ,"
        sStr = sStr & "        K12 AS " & constSatams(8) & "_��       ,"
        sStr = sStr & "        J13 AS " & constSatams(9) & "          ,"
        sStr = sStr & "        K13 AS " & constSatams(9) & "_��       ,"
        sStr = sStr & " '' AS J14, "
        sStr = sStr & " '' AS K14, "
                                           
        sStr = sStr & "        J15 AS ����_����     ,"
        sStr = sStr & "        K15 AS ����_����_��  ,"
        sStr = sStr & "        J16 AS �Ͼ�_�̻�     ,"
        sStr = sStr & "        K16 AS �Ͼ�_�̻�_��  ,"
        sStr = sStr & "        J17 AS ����_Ȯ��     ,"
        sStr = sStr & "        K17 AS ����_Ȯ��_��  ,"
        sStr = sStr & "        J18 AS �Ҿ�_������   ,"
        sStr = sStr & "        K18 AS �Ҿ�_������_��,"
                                   
        sStr = sStr & "        J19 AS �߾�          ,"
        sStr = sStr & "        K19 AS �߾�_��       ,"
        sStr = sStr & "        J20 AS �ѹ�          ,"
        sStr = sStr & "        K20 AS �ѹ�_��       ,"
        sStr = sStr & "        J21 AS �ƶ���        ,"
        sStr = sStr & "        K21 AS �ƶ���_��     ,"
        
        ' �뷮�� ��û�� ���� �����ܴ�... �׷��� �ѿ����� ������ �߰��س����̴�.. �׷��� ������ �ȵǾ�����.
        ' �ؿ��ٰ���.. ����..
        sStr = sStr & "        D_UNIVCD AS ��������, D_MAJORCD AS �����ܴ� "
        
        sStr = sStr & "    FROM ( "
    
            sStr = sStr & "  SELECT A.SCHNO           ,"
            sStr = sStr & "         MAX(ACID      ) AS ACID       ,"
            sStr = sStr & "         MAX(EXMID     ) AS EXMID      ,"
            sStr = sStr & "         MAX(STDNM     ) AS STDNM      ,"
            sStr = sStr & "         MAX(D_UNIVCD     ) AS D_UNIVCD      ,"
            sStr = sStr & "         MAX(D_MAJORCD     ) AS D_MAJORCD      ,"
            sStr = sStr & "         MAX(birth_ymd     ) AS birth_ymd      ,"
            sStr = sStr & "         MAX(EXMTYPE   ) AS EXMTYPE    , MAX(KAEYOL    ) AS KAEYOL     ,"
            sStr = sStr & "         MAX(SEL1      ) AS SEL1       , MAX(SEL2      ) AS SEL2       , MAX(SEL3      ) AS SEL3      , MAX(SEL4      ) AS SEL4      , MAX(SEL5      ) AS  SEL5      ,"
            sStr = sStr & "         MAX(K_NUM     ) AS K_NUM      , MAX(M_NUM     ) AS M_NUM      , MAX(E_NUM     ) AS E_NUM     , MAX(TOT_NUM   ) AS TOT_NUM   , MAX(N_NUM   ) AS N_NUM   ,"
            sStr = sStr & "         MAX(SEL1_SCH  ) AS SEL1_SCH   , MAX(SEL2_SCH  ) AS SEL2_SCH   ,"
            sStr = sStr & "         MAX(PASS1     ) AS PASS1      , MAX(PASS2     ) AS PASS2      , MAX(PASS3     ) AS PASS3     , MAX(PASS4     ) AS PASS4     , MAX(CL_CLOSE  ) AS  CL_CLOSE  ,"
            sStr = sStr & "         MAX(CY_ACNT   ) AS CY_ACNT    , MAX(TOT_AMT   ) AS TOT_AMT    ,"
            sStr = sStr & "         MAX(BASE_AMT1 ) AS BASE_AMT1  , MAX(BASE_AMT2 ) AS BASE_AMT2  , MAX(BASE_AMT3 ) AS BASE_AMT3 , MAX(BASE_AMT4 ) AS BASE_AMT4 ,"
            sStr = sStr & "         MAX(BASE_AMT5 ) AS BASE_AMT5  , MAX(BASE_AMT6 ) AS BASE_AMT6  , MAX(BASE_AMT7 ) AS BASE_AMT7 , MAX(BASE_AMT8 ) AS BASE_AMT8 ,"
            sStr = sStr & "         MAX(TAMGU_AMT1) AS TAMGU_AMT1 , MAX(TAMGU_AMT2) AS TAMGU_AMT2 , MAX(TAMGU_AMT3) AS TAMGU_AMT3, MAX(TAMGU_AMT4) AS TAMGU_AMT4, MAX(TAMGU_AMT5) AS  TAMGU_AMT5,"
            sStr = sStr & "         MAX(TAMGU_AMT6) AS TAMGU_AMT6 , MAX(TAMGU_AMT7) AS TAMGU_AMT7 , MAX(TAMGU_AMT8) AS TAMGU_AMT8, MAX(TAMGU_AMT9) AS TAMGU_AMT9, MAX(TAMGU_AMT10) AS TAMGU_AMT10, MAX(TAMGU_AMT11) AS TAMGU_AMT11,"
            sStr = sStr & "         MAX(SEX       ) AS SEX        ,"
            sStr = sStr & "         MAX(ZIP       ) AS ZIP        , MAX(ADDR1     ) AS ADDR1      , MAX(ADDR2     ) AS ADDR2     ,"
            sStr = sStr & "         MAX(TEL       ) AS TEL        , MAX(CEL       ) AS CEL        , MAX(EMAIL     ) AS EMAIL     ,"
            sStr = sStr & "         MAX(HIGH_SCH  ) AS HIGH_SCH   , MAX(GRADE_YEAR) AS GRADE_YEAR ,"
            sStr = sStr & "         MAX(PRNT_NM   ) AS PRNT_NM    , MAX(PRNT_RLTN ) AS PRNT_RLTN  ,"
            sStr = sStr & "         MAX(PRNT_ZIP  ) AS PRNT_ZIP   , MAX(PRNT_ADDR1) AS PRNT_ADDR1 , MAX(PRNT_ADDR2) AS PRNT_ADDR2,"
            sStr = sStr & "         MAX(PRNT_TEL  ) AS PRNT_TEL   , MAX(PRNT_CEL  ) AS PRNT_CEL   , MAX(PRNT_JOB  ) AS PRNT_JOB  , MAX(PRNT_W_TEL) AS PRNT_W_TEL,"
            sStr = sStr & "         MAX(PHOTO_PATH) AS PHOTO_PATH , MAX(R_WAY     ) AS R_WAY      , MAX(ORD_NO    ) AS ORD_NO    , "
            sStr = sStr & "         MAX(TO_CHAR(REGDATE,'YYYY-MM-DD HH24:MI:SS')) AS REGDATE      , MAX(PTS_SEL   ) AS PTS_SEL   , MAX(MU_TYPE   ) AS MU_TYPE   , MAX(CASH_BILL_NUM) AS CASH_BILL_NUM"
            
                    sStr = sStr & " , "
                    sStr = sStr & "        SUM(J01) AS J01,"
                    sStr = sStr & "        SUM(K01) AS K01,"
                    sStr = sStr & "        SUM(J02) AS J02,"
                    sStr = sStr & "        SUM(K02) AS K02,"
                    sStr = sStr & "        SUM(J03) AS J03,"
                    sStr = sStr & "        SUM(K03) AS K03,"
                    
                    sStr = sStr & "        SUM(J04) AS J04,"
                    sStr = sStr & "        SUM(K04) AS K04,"
                    sStr = sStr & "        SUM(J05) AS J05,"
                    sStr = sStr & "        SUM(K05) AS K05,"
                    sStr = sStr & "        SUM(J06) AS J06,"
                    sStr = sStr & "        SUM(K06) AS K06,"
                    sStr = sStr & "        SUM(J07) AS J07,"
                    sStr = sStr & "        SUM(K07) AS K07,"
                    sStr = sStr & "        SUM(J08) AS J08,"
                    sStr = sStr & "        SUM(K08) AS K08,"
                    sStr = sStr & "        SUM(J09) AS J09,"
                    sStr = sStr & "        SUM(K09) AS K09,"
                    sStr = sStr & "        SUM(J10) AS J10,"
                    sStr = sStr & "        SUM(K10) AS K10,"
                    sStr = sStr & "        SUM(J11) AS J11,"
                    sStr = sStr & "        SUM(K11) AS K11,"
                    
                    sStr = sStr & "        SUM(J12) AS J12,"
                    sStr = sStr & "        SUM(K12) AS K12,"
                    sStr = sStr & "        SUM(J13) AS J13,"
                    sStr = sStr & "        SUM(K13) AS K13,"
                    sStr = sStr & " '' AS K14,"
                    sStr = sStr & " '' AS J14,"
                    
                    sStr = sStr & "        SUM(J15) AS J15,"
                    sStr = sStr & "        SUM(K15) AS K15,"
                    sStr = sStr & "        SUM(J16) AS J16,"
                    sStr = sStr & "        SUM(K16) AS K16,"
                    sStr = sStr & "        SUM(J17) AS J17,"
                    sStr = sStr & "        SUM(K17) AS K17,"
                    sStr = sStr & "        SUM(J18) AS J18,"
                    sStr = sStr & "        SUM(K18) AS K18,"
                    
                    sStr = sStr & "        SUM(J19) AS J19,"
                    sStr = sStr & "        SUM(K19) AS K19,"
                    sStr = sStr & "        SUM(J20) AS J20,"
                    sStr = sStr & "        SUM(K20) AS K20,"
                    sStr = sStr & "        SUM(J21) AS J21,"
                    sStr = sStr & "        SUM(K21) AS K21"
            
            sStr = sStr & "    FROM ("
            '---------------------------------------------------------------------------- ��ü�л� ��ȸ START
            sStr = sStr & "          SELECT *"
            sStr = sStr & "            FROM CLSTD01TB"
            sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "             AND EXMID > ' ' "
            
    If Trim(Right(kaeyol, 30)) <> "ALL" Then
            sStr = sStr & "             AND KAEYOL = '" & Trim(Right(kaeyol, 30)) & "'"
    End If
    
    '<< �Ⱓ���� >>
    If day1 <> "" Or day2 <> "" Then
        If Trim(day1) <> "" And Trim(day2) <> "" Then
            sStr = sStr & "             AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                             AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) = "" And Trim(day2) <> "" Then
            sStr = sStr & "             AND REGDATE BETWEEN TO_DATE('19000101000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                             AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) <> "" And Trim(day2) = "" Then
            sStr = sStr & "             AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                             AND TO_DATE('99991231235959', 'YYYYMMDDHH24MISS') "
        End If

    End If
            sStr = sStr & "             AND BIGO2 IS NULL "
            sStr = sStr & "          UNION ALL"
            '---------------------------------------------------------------------------- ��ü�л� ��ȸ END
            '---------------------------------------------------------------------------- �հ��� ��ȸ START
            sStr = sStr & "          SELECT *"
            sStr = sStr & "            From CLSTD01TB"
            sStr = sStr & "           WHERE (PASS1 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS2 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS3 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS4 = '" & Trim(basModule.SchCD) & "'" & " )"
            sStr = sStr & "             AND EXMID > ' ' "
    If Trim(Right(kaeyol, 30)) <> "ALL" Then
            sStr = sStr & "             AND KAEYOL = '" & Trim(Right(kaeyol, 30)) & "'"
    End If
    
    '<< �Ⱓ���� >>
    If day1 <> "" Or day2 <> "" Then
        If Trim(day1) <> "" And Trim(day2) <> "" Then
            sStr = sStr & "             AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                             AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) = "" And Trim(day2) <> "" Then
            sStr = sStr & "             AND REGDATE BETWEEN TO_DATE('19000101000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                             AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) <> "" And Trim(day2) = "" Then
            sStr = sStr & "             AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                             AND TO_DATE('99991231235959', 'YYYYMMDDHH24MISS') "
        End If

    End If
            sStr = sStr & "             AND BIGO2 IS NULL "
            
            sStr = sStr & "          ) A, "
            
            sStr = sStr & "               ("
        
                sStr = sStr & "         SELECT SCHNO,"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J01,    /* ���                  */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K01,    /* �����  ���          */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J02,    /* ��������              */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K02,    /* �����  ��������      */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J03,    /* �ܱ���                */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K03,    /* �����  �ܱ���        */"
                
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(0) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '51', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J04,    /* ��Ž-" & constSatams(0) & "       , ��Ž-����1             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(0) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '51', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K04,    /* �����  ��Ž-" & constSatams(0) & "        , ��Ž-����1     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(1) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '52', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J05,    /* ��Ž-" & constSatams(1) & "        , ��Ž-ȭ��1             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(1) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '52', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K05,    /* �����  ��Ž-" & constSatams(1) & "        , ��Ž-ȭ��1     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '53', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J06,    /* ��Ž-" & constSatams(2) & "        , ��Ž-�������1             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '53', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K06,    /* �����  ��Ž-" & constSatams(2) & "        , ��Ž-�������1     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '54', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J07,    /* ��Ž-" & constSatams(3) & "  , ��Ž-��������1         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '54', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K07,    /* �����  ��Ž-" & constSatams(3) & "  , ��Ž-��������1 */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '55', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J08,    /* ��Ž-" & constSatams(4) & "      , ��Ž-����2             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '55', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K08,    /* �����  ��Ž-" & constSatams(4) & "      , ��Ž-����2     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '56', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J09,    /* ��Ž-" & constSatams(5) & "    , ��Ž-ȭ��2             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '56', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K09,    /* �����  ��Ž-" & constSatams(5) & "    , ��Ž-ȭ��2     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '57', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J10,      /* ��Ž-" & constSatams(6) & "    , ��Ž-�������2           */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '57', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K10,      /* ����� ��Ž-" & constSatams(6) & "    , ��Ž-�������2    */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(7) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '58', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J11,    /* ��Ž-" & constSatams(7) & "        , ��Ž-��������2         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(7) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '58', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K11,    /* �����  ��Ž-" & constSatams(7) & "        , ��Ž-��������2 */"
                
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(8) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J12,    /* ��Ž-" & constSatams(8) & "         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(8) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K12,    /* �����  ��Ž-" & constSatams(8) & " */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(9) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J13,    /* ��Ž-" & constSatams(9) & "         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(9) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K13,    /* �����  ��Ž-" & constSatams(9) & " */"
                sStr = sStr & " '' AS K14,"
                sStr = sStr & " '' AS J14,"
                
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_NUM,'X',0, SUB_NUM), '81', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J15,    /* ����             , ������                 */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_BAK,'X',0, SUB_BAK), '81', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K15,    /* �����  ����             , ������         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_NUM,'X',0, SUB_NUM), '82', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J16,    /* �Ͼ�             , �̻����               */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_BAK,'X',0, SUB_BAK), '82', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K16,    /* �����  �Ͼ�             , �̻����       */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_NUM,'X',0, SUB_NUM), '83', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J17,    /* �����ĳ�         , Ȯ�����               */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_BAK,'X',0, SUB_BAK), '83', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K17,    /* �����  �����ĳ�         , Ȯ�����       */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_NUM,'X',0, SUB_NUM), '43', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J18,    /* �Ҿ�             , ��������               */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_BAK,'X',0, SUB_BAK), '43', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K18,    /* �����  �Ҿ�             , ��������       */"
                
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J19,    /* �߱���                */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K19,    /* �����  �߱���        */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J20,    /* �ѹ�                  */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K20,    /* �����  �ѹ�          */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J21,    /* �ƶ���                */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K21     /* �����  �ƶ���        */"
                sStr = sStr & "           FROM CLSTD03TB"
        
        sStr = sStr & "                ) B"
        sStr = sStr & "        WHERE A.SCHNO = B.SCHNO(+)"
            
            sStr = sStr & "   GROUP BY A.SCHNO"
            '---------------------------------------------------------------------------- �հ��� ��ȸ END
    
    sStr = sStr & "    ) "
    sStr = sStr & " ORDER BY EXMID "
    
    Get_StdExcuteSqlToExcel = sStr
End Function


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'��ƿ
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Public Function Cbo_Val(ByRef cboControl As Object, ByVal length As Long) As String
    Cbo_Val = Trim(Right(cboControl.Text, 30))
End Function


Public Function Get_IndexByChk(ByRef Control As Object) As Long
    
    Dim i As Long
    Dim selIndex As Long
    
    selIndex = -1
    For i = 0 To Control.count - 1
        If Control(i).value = True Then: selIndex = i
    Next
    
    Get_IndexByChk = selIndex
End Function
