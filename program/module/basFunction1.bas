Attribute VB_Name = "basGwamok"
Option Explicit


Public SATAM_COUNT
Public GWATAM_COUNT
Public ENG2_COUNT
Public MATH2_COUNT
Public TAMGOO_COUNT
Public CLINIC_L_COUNT
Public CLINIC_M_COUNT
Public CLINIC_E_COUNT

'>> ��� �̴�..
Public Const CLINIC_L_CLASS = 101
Public Const CLINIC_M_CLASS = 111
Public Const CLINIC_E_CLASS = 121
'>> �Ʒ� ����� �߸������ߴ�. ���� 21,31,51������ �Ǿ�����ߴµ�.. ������.
Public Const SATAM_CLASS = 20
Public Const GWATAM_CLASS = 30
Public Const ENG2_CLASS = 50
Public Const MATH2_CLASS = 80
Public Const TAMGOO_CLASS = 90

Public constSatams() As String
Public constGwatams() As String
Public constEng2s() As String
Public constMaths() As String
Public constTamgoos() As String

Public constSatamCodes() As String
Public constGwatamCodes() As String
Public constEng2Codes() As String
Public constMathCodes() As String
Public constTamgooCodes() As String


Public g_sClinic_LCodes() As String
Public g_sClinic_MCodes() As String
Public g_sClinic_ECodes() As String

Public g_sClinic_Ls() As String
Public g_sClinic_Ms() As String
Public g_sClinic_Es() As String

Public Const TGANG_CODE = "95"





'���� : ����� �ڵ尡 ������ �̷��������. basGwamok���ϸ� �����ϸ� �ǵ���.
'>>     *********�߿� ************** 1. frm���Ͽ��� �������� ��� �Ⱦ��.
'>>     100%�Ⱦ�����Ѵ�. (���� ->�����ĳ�, ����->�����)


'��Ȳ : �ڷ������ �����ϱ������� DBó�� ����ϸ� �ǳ�
'>>     �ҽ��ڵ带 ���� uExcel_StdData.SATAM3(�������Ͻ�Ʈ��SATAM3�÷�)������ �ϵ��ڵ��Ǿ��մºκ��� �ֱ⋚����
'>>     ��¿������ ������ ���Ǵºκ� ��ü�� �߻��캸�ƾ� ������
'>>     �����ؾ��ϴºκ��� ���� �ٿ��־���.

'>>     ���α׷� �ҽ��ڵ��� ���⼺�� ����ϹǷ�.. private ���� �Ǿ��ִ°��� public���� �ٲپ ���к��ϰ�
'>>     �ٸ����Ͽ��� ������� �ʵ����Ѵ�.
'>>     ��ü�����丵�� ���� �ʴ��̻�... ���̻� �ҽ����������� �ʵ��� ��������.

'>>     �п����� ������ Ʋ���� ������ �п� �ڵ忡 ���� �迭���� �������־���Ѵ�.

'DB��ſ� �迭�� ���.
'�������� DB�� �̿��Ͽ� ���������� �̿��Ҽ��ֶǷ�
'���񺯰��ϸ鼭 �ҽ������� ������.



'>>     ���ʿ����. �����ݰ� �κ������� �����ϴٰ� �� ����������. (�ڵ��� �ϰ��� �ı�)
'>>     �밡�� �Ұ� �ϰ�. �׳� �ִ��� ������ �߽����θ� �ٲٷ��� ����ߴ�.


'����
Private Function setConstant_M()
    
End Function
'����
Private Function setConstant_P()
    
End Function
'�λ�
Private Function setConstant_B()
   
End Function
'����
Private Function setConstant_S()
    '>>Ŭ����
    '����
    g_sClinic_Ls(0) = "(��ȭ)���&���� �����"
    g_sClinic_Ls(1) = "(��ȭ)���� ������� ����"
    g_sClinic_Ls(2) = "(�⺻)���&���� �����"
    g_sClinic_Ls(3) = "(��ȭ)���� ������� ����"
    
    g_sClinic_LCodes(0) = "101"
    g_sClinic_LCodes(1) = "102"
    g_sClinic_LCodes(2) = "103"
    g_sClinic_LCodes(3) = "104"
    
    '����
    g_sClinic_Ms(0) = "(�⺻)���ɿ� �� �ʿ��� ��1 ����"
    g_sClinic_Ms(1) = "(��ȭ)������ �Լ�"
    g_sClinic_Ms(2) = "(�⺻)�Լ��� ����&�̺�"
    g_sClinic_Ms(3) = "(��ȭ)���������� ����"
    
    g_sClinic_MCodes(0) = "111"
    g_sClinic_MCodes(1) = "112"
    g_sClinic_MCodes(2) = "113"
    g_sClinic_MCodes(3) = "114"
    
    '����
    g_sClinic_Es(0) = "(�⺻)�ٽɹ����� ����"
    g_sClinic_Es(1) = "(��ȭ)������� ���ؿ���"
    g_sClinic_Es(2) = "(�⺻)���ʹ��� �� ����"
    g_sClinic_Es(3) = "(��ȭ)������� ���ؿ���"
    
    g_sClinic_ECodes(0) = "121"
    g_sClinic_ECodes(1) = "122"
    g_sClinic_ECodes(2) = "123"
    g_sClinic_ECodes(3) = "124"
End Function
'�뷮��
Private Function setConstant_N()
    '>>Ŭ����
    '����
    g_sClinic_Ls(0) = "(��ȭ)���� ����"
    g_sClinic_Ls(1) = "(��ȭ)���� ����"
    g_sClinic_Ls(2) = "(�⺻)���� ������ ����"
    g_sClinic_Ls(3) = "(�⺻)���� ������ ����"
    
    g_sClinic_LCodes(0) = "101"
    g_sClinic_LCodes(1) = "102"
    g_sClinic_LCodes(2) = "103"
    g_sClinic_LCodes(3) = "104"
    
    '����
    g_sClinic_Ms(0) = "(��ȭ)������ �Լ�"
    g_sClinic_Ms(1) = "(�⺻)���ɿ� �� �ʿ��� ��1 ����"
    g_sClinic_Ms(2) = "(��ȭ)���������� ����"
    g_sClinic_Ms(3) = "(�⺻)����1 �ٽ����� ������"
    
    g_sClinic_MCodes(0) = "111"
    g_sClinic_MCodes(1) = "112"
    g_sClinic_MCodes(2) = "113"
    g_sClinic_MCodes(3) = "114"
    
    '����
    g_sClinic_Es(0) = "(��ȭ)���� ����&����"
    g_sClinic_Es(1) = "(��ȭ)���"
    g_sClinic_Es(2) = "(�⺻)���� ���� Ǯ�̹�"
    g_sClinic_Es(3) = "(�⺻)���"
    
    g_sClinic_ECodes(0) = "121"
    g_sClinic_ECodes(1) = "122"
    g_sClinic_ECodes(2) = "123"
    g_sClinic_ECodes(3) = "124"
End Function
'����
Private Function setConstant_K()

End Function
'�߹�
Private Function setConstant_Q()

End Function

'�ֹ�
Private Function setConstant_W()

End Function

'����
Private Function setConstant_J()
    
'    SATAM_COUNT = 11
'    GWATAM_COUNT = 9
'
'
'    ReDim Preserve constSatams(SATAM_COUNT - 1)
'    ReDim Preserve constSatamCodes(SATAM_COUNT - 1)
'    ReDim Preserve constGwatams(GWATAM_COUNT - 1)
'    ReDim Preserve constGwatamCodes(GWATAM_COUNT - 1)
'
'
'
'    constSatams(10) = "Ư��"
'    constSatamCodes(10) = "95"
'    constGwatams(8) = "Ư��"
'    constGwatamCodes(8) = "95"
    
End Function



'DB�迭
Function setConstant()

    SATAM_COUNT = 10
    GWATAM_COUNT = 8
    ENG2_COUNT = 12
    MATH2_COUNT = 4
    TAMGOO_COUNT = 3
    CLINIC_L_COUNT = 4
    CLINIC_M_COUNT = 4
    CLINIC_E_COUNT = 4
    
    
    ReDim constSatams(SATAM_COUNT - 1)
    ReDim constGwatams(GWATAM_COUNT - 1)
    ReDim constEng2s(ENG2_COUNT - 1)
    ReDim constMaths(MATH2_COUNT - 1)
    ReDim constTamgoos(TAMGOO_COUNT - 1)
    
    ReDim constSatamCodes(SATAM_COUNT - 1)
    ReDim constGwatamCodes(GWATAM_COUNT - 1)
    ReDim constEng2Codes(ENG2_COUNT - 1)
    ReDim constMathCodes(MATH2_COUNT - 1)
    ReDim constTamgooCodes(TAMGOO_COUNT - 1)

    
    ReDim g_sClinic_LCodes(CLINIC_L_COUNT - 1)
    ReDim g_sClinic_MCodes(CLINIC_M_COUNT - 1)
    ReDim g_sClinic_ECodes(CLINIC_E_COUNT - 1)
    
    ReDim g_sClinic_Ls(CLINIC_E_COUNT - 1)
    ReDim g_sClinic_Ms(CLINIC_E_COUNT - 1)
    ReDim g_sClinic_Es(CLINIC_E_COUNT - 1)



    ' ��ȸŽ��
    constSatams(0) = "�ѱ���"
    constSatams(1) = "�����"
    constSatams(2) = "���ƽþƻ�"
    constSatams(3) = "�ѱ�����"
    constSatams(4) = "��������"
    constSatams(5) = "��Ȱ������"
    constSatams(6) = "�������"
    constSatams(7) = "������ġ"
    constSatams(8) = "����"
    constSatams(9) = "��ȸ��ȭ"
    
    
    constSatamCodes(0) = "21"
    constSatamCodes(1) = "22"
    constSatamCodes(2) = "23"
    constSatamCodes(3) = "24"
    constSatamCodes(4) = "25"
    constSatamCodes(5) = "26"
    constSatamCodes(6) = "27"
    constSatamCodes(7) = "28"
    constSatamCodes(8) = "29"
    constSatamCodes(9) = "30"
    
    '����Ž��
    constGwatams(0) = "����1"
    constGwatams(1) = "ȭ��1"
    constGwatams(2) = "�������1"
    constGwatams(3) = "��������1"
    constGwatams(4) = "����2"
    constGwatams(5) = "ȭ��2"
    constGwatams(6) = "�������2"
    constGwatams(7) = "��������2"
    
    constGwatamCodes(0) = "51"
    constGwatamCodes(1) = "52"
    constGwatamCodes(2) = "53"
    constGwatamCodes(3) = "54"
    constGwatamCodes(4) = "55"
    constGwatamCodes(5) = "56"
    constGwatamCodes(6) = "57"
    constGwatamCodes(7) = "58"
    
    ' ��2�ܱ���
    constEng2s(0) = "����"
    constEng2s(1) = "�Ͼ�"
    constEng2s(2) = "�����ĳľ�"
    constEng2s(3) = "�Ҿ�"
    constEng2s(4) = "�߱���"
    constEng2s(5) = "�ѹ�"
    constEng2s(6) = "���"
    constEng2s(7) = "����"
    constEng2s(8) = "����"
    constEng2s(9) = "�����"
    constEng2s(10) = "��������"
    constEng2s(11) = "�ƶ���"
    
     
    constEng2Codes(0) = "31"
    constEng2Codes(1) = "32"
    constEng2Codes(2) = "33"
    constEng2Codes(3) = "34"
    constEng2Codes(4) = "35"
    constEng2Codes(5) = "36"
    constEng2Codes(6) = "37"
    constEng2Codes(7) = "38"
    constEng2Codes(8) = "39"
    constEng2Codes(9) = "40"
    constEng2Codes(10) = "41"
    constEng2Codes(11) = "42"
    
    'Ž������
    constMaths(0) = "������"
    constMaths(1) = "�̻����"
    constMaths(2) = "Ȯ�����"
    constMaths(3) = "��������"
    
    constMathCodes(0) = "81"
    constMathCodes(1) = "82"
    constMathCodes(2) = "83"
    constMathCodes(3) = "84"
    
    
    'Ž������
    constTamgoos(0) = "���"
    constTamgoos(1) = "����"
    constTamgoos(2) = "�ܱ���"
    
    constTamgooCodes(0) = "91"
    constTamgooCodes(1) = "92"
    constTamgooCodes(2) = "93"
    
    

    
    
    '���� �ڵ���� �⺻�̰� ������ �ٲ�°͵��� �Ʒ��� �п��� ���� �Լ����� ������.
    Select Case basModule.SchCD
        Case "M":   Call setConstant_M
        Case "P":   Call setConstant_P
        Case "B":   Call setConstant_B
        Case "J":   Call setConstant_J
        Case "S":   Call setConstant_S
        Case "N":   Call setConstant_N
        Case "K":   Call setConstant_K
        Case "Q":   Call setConstant_Q
        Case "W":   Call setConstant_W
    End Select
    
    
End Function

'����� - > �����ڵ�
Function Get_GwaMokCodeByName(subject As String) As String
    
    Dim i As Integer
    
    For i = 0 To SATAM_COUNT - 1
        If constSatams(i) = subject Then
            Get_GwaMokCodeByName = constSatamCodes(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To ENG2_COUNT - 1
        If constEng2Codes(i) = subject Then
            Get_GwaMokCodeByName = constEng2Codes(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To GWATAM_COUNT - 1
        If constGwatams(i) = subject Then
            Get_GwaMokCodeByName = constGwatamCodes(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To MATH2_COUNT - 1
        If constMaths(i) = subject Then
            Get_GwaMokCodeByName = constMathCodes(i)
            Exit Function
        End If
    Next i
    For i = 0 To TAMGOO_COUNT - 1
        If constTamgoos(i) = subject Then
            Get_GwaMokCodeByName = constTamgooCodes(i)
            Exit Function
        End If
    Next i

End Function



'�ʵ忡 �ش��ϴ� �����ڵ��ϵ��� �������ش�.
'SEL1~5�� DB �÷���.
Function Get_GwaMokCodes(fieldName As String) As String()

    Select Case fieldName
        Case "SEL1": Get_GwaMokCodes = constSatamCodes  '��Ž
        Case "SEL2": Get_GwaMokCodes = constEng2Codes   '��2�ܱ���
        Case "SEL3": Get_GwaMokCodes = constGwatamCodes '��Ž
        Case "SEL4": Get_GwaMokCodes = constMathCodes   '����
        Case "SEL5": Get_GwaMokCodes = constTamgooCodes 'Ž��
        
    End Select

End Function

'�ʵ忡 �ش��ϴ� �����ڵ��ϵ��� �������ش�.
Function Get_GwaMokNames(fieldName As String) As String()

    Select Case fieldName
        Case "SEL1": Get_GwaMokNames = constSatams
        Case "SEL2": Get_GwaMokNames = constEng2s
        Case "SEL3": Get_GwaMokNames = constGwatams
        Case "SEL4": Get_GwaMokNames = constMaths
        Case "SEL5": Get_GwaMokNames = constTamgoos
        
    End Select

End Function

' ���� �ȿ� ������� �������ش�.
Function Get_StrGwaMokRange(codes As String, rangeStart As Long, rangeEnd As Long) As String

    'codes�� split
    Dim arrTmp() As String
    Dim count As Long
    Dim i As Long
    Dim code As String
    
    Dim sReturnVal As String
    
    
    sReturnVal = ""
    
    arrTmp = Split(Trim(codes), "|", -1, vbTextCompare)
    
    count = UBound(arrTmp)
    
    For i = 0 To count - 1
    
        code = arrTmp(i)
        If code >= rangeStart And code <= rangeEnd Then
        
            sReturnVal = sReturnVal & Get_StrGwaMokByCode(code)
        
        End If
        
    Next i
    
    Get_StrGwaMokRange = sReturnVal

End Function

'���� �ڵ� ->  �����
Function Get_StrGwaMokByCode(gwamokCode As String) As String

   Dim i As Integer
    
    For i = 0 To SATAM_COUNT - 1
        If constSatamCodes(i) = gwamokCode Then
            Get_StrGwaMokByCode = constSatams(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To ENG2_COUNT - 1
        If constEng2Codes(i) = gwamokCode Then
            Get_StrGwaMokByCode = constEng2s(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To GWATAM_COUNT - 1
        If constGwatamCodes(i) = gwamokCode Then
            Get_StrGwaMokByCode = constGwatams(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To MATH2_COUNT - 1
        If constMathCodes(i) = gwamokCode Then
            Get_StrGwaMokByCode = constMaths(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To TAMGOO_COUNT - 1
        If constTamgooCodes(i) = gwamokCode Then
            Get_StrGwaMokByCode = constTamgoos(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To CLINIC_L_COUNT - 1
        If g_sClinic_LCodes(i) = gwamokCode Then
            Get_StrGwaMokByCode = g_sClinic_Ls(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To CLINIC_M_COUNT - 1
        If g_sClinic_MCodes(i) = gwamokCode Then
            Get_StrGwaMokByCode = g_sClinic_Ms(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To CLINIC_E_COUNT - 1
        If g_sClinic_ECodes(i) = gwamokCode Then
            Get_StrGwaMokByCode = g_sClinic_Es(i)
            Exit Function
        End If
    Next i
    
End Function

'�������忡�� ����11��° �÷��� �����Ѵ�.
'������ ��忡�� ����11 �÷��� ���� �����ϴ¹���� �𸣰ڴ�.
Public Sub Spread_DelCel(ss As Control, colNum As Long)
    ss.Col = colNum
    ss.Action = 4
End Sub
