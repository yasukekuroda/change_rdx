Attribute VB_Name = "����"
Option Explicit
'---------------------------------------------------------------------------
'����
'  �ϊ����������ł��Ă��邩���Ғl�Ƃ��ĂĔ�r����B
'
'���͔͈�                                     ����
'   2�i�� : 0�`11_1111_1111_1111_1111_1111  ||  OK
'  10�i�� : 0�`4,194,303                    ||  OK
'  16�i�� : 0�`3F_FFFF                      ||  OK
'
'���s����
'  1���  31��20�b
'  2���  24��40�b
'  3���  19��40�b
'---------------------------------------------------------------------------
Public Sub INSPECTION()
  Dim i As Long
  Call �}�N���������J�n

  For i = 0 To 4194303 '���̐����̑S���̓p�^�[��
  
    ' 2�i����10�i��, 16�i��
    '(A���@�ŕϊ��������ʂ������ɂ��邪�AMain�֐����ōēxAB�ɕϊ�����r����)
    Call Main������.Main(True, True, "2�i��", ��ϊ�_A.DecToBin(str(i)))
    
    '10�i���� 2�i��, 16�i��
    Call Main������.Main(True, True, "10�i��", i)
    
    '16�i���� 2�i��, 10�i��
    Call Main������.Main(True, True, "16�i��", Hex(i_dat))
  Next i
    
  Call �}�N���������I��
  MsgBox "���s����"
End Sub


'---------------------------------------------------------------------------
'�l��r
'  ����1   (IN)    �FA���@�̌���
'  ����2   (IN)    �FB���@�̌���
'  �߂�l  (OUT)   �F�Ȃ�
'---------------------------------------------------------------------------
Public Sub COMPARE(ByVal str1 As String, _
                   ByVal str2 As String)
  If str1 <> str2 Then
    MsgBox "�l�s��v�ł��B" & vbCrLf & _
           "str1 : " & str1 & vbCrLf & _
           "str2 : " & str2
  End If
End Sub
                
                
'##########################################################
'�}�N���������ƌx����~��
'##########################################################
Private Sub �}�N���������J�n()

Application.Interactive = False    '�L�[�{�[�h�̓���OFF
Application.ScreenUpdating = False '��ʕ`�ʂ��~
Application.Cursor = xlWait        '�J�[�\���������v�^��
Application.EnableEvents = False   '�C�x���g��}�~
Application.DisplayAlerts = False  '�x�����b�Z�[�W��\��
Application.Calculation = xlCalculationManual '�v�Z���蓮��

End Sub


Private Sub �}�N���������I��()

'Application.StatusBar = False    '�X�e�[�^�X�o�[��\��
Application.Calculation = xlCalculationAutomatic '�v�Z������
Application.DisplayAlerts = True  '���b�Z�[�W�\���J�n
Application.EnableEvents = True   '�C�x���g����t�J�n
Application.Cursor = xlDefault    '�J�[�\�����f�t�H���g
Application.ScreenUpdating = True '��ʕ`��J�n
Application.Interactive = True    '�L�[�{�[�h���͎�t�J�n

End Sub


'  '�f�o�b�N : �z��̒l�\��
'  Dim i As Long
'  For i = LBound(ARR_STR, 1) To UBound(ARR_STR, 1)
'    If ARR_STR(i, 0) <> "" Then
'      MsgBox "ARR_STR[" & i & "][0] = """ & ARR_STR(i, 0) & """   " & _
'             "ARR_STR[" & i & "][1] = """ & ARR_STR(i, 1) & """   " & _
'             "ARR_STR[" & i & "][2] = """ & ARR_STR(i, 2) & """   " & _
'             "ARR_CNT[" & i & "] = " & ARR_CNT(i) & "��"
'    End If
'  Next i





