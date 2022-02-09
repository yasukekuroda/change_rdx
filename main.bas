Attribute VB_Name = "Main������"
Option Explicit
'###########################################################################
'��ϊ��X�N���v�g
'2022/01/31   ���c
'�����̗L�� / �����_�̗L��  �ɖ��Ή��ł��B�B
'
'���p
'  VBA�ŕ�����̑S�p�p�����𔼊p�ɕϊ�����
'  https://vbabeginner.net/convert-full-width-alphanumeric-characters-in-a-character-string-to-half-width/
'###########################################################################
'---------------------------------------------------------------------------
'�z���`  (�{���W���[�����̂ݗL��)
'---------------------------------------------------------------------------
'�ϊ���̕�������i�[ (xx,0)��2�i��, (xx,1)��:10�i��, (xx,2)��:16�i��
Dim ARR_STR(ARR_MAX - 1, 2) As String

'��ʕ`�ʗp�B���I�B("_"��","��}��)
Dim ARR_STR_DISP() As String

'�o���񐔂��J�E���g�B
'2�i����"111"�ƌ����������̂�10�i����"7"�ƌ����������͓̂���Ƃ��Ĉ����B
Dim ARR_CNT(ARR_MAX - 1) As Long


'###########################################################################
'
'Main �֐�
'
'���ؗp
'Public Sub Main(ByVal i_pls As Boolean, _
'                ByVal i_int As Boolean, _
'                ByVal i_rdx As String, _
'                ByVal i_dat As String)
'###########################################################################
Public Sub Main()

  '-------------------------------------------------------------------------
  '�ϐ��錾
  '-------------------------------------------------------------------------
  Dim i_pls As Boolean     'Tlue:���̐��̂�,  False:���̐����܂�
  Dim i_int As Boolean     'Tlue:����,        False:�������܂�
  Dim i_rdx As String      '�ϊ��O�̊("2�i��/10�i��/16�i��")
  Dim i_dat As String      '���͕�����
  '�ݒ�l�ǂݍ���
  i_pls = Pg_I_PLS.Value = "�Ȃ�"
  i_int = Pg_I_INT.Value = "�Ȃ�"
  i_rdx = Pg_I_RDX.Value
  i_dat = Pg_I_DAT.Value

  '�ϊ���̕�����i�[
  Dim o_radix__2 As String ' 2�i��
  Dim o_radix_10 As String '10�i��
  Dim o_radix_16 As String '16�i��

  '-------------------------------------------------------------------------
  '���͒l�̘_���`�F�b�N
  '  �߂�l�� False �̂Ƃ��A�}�N�����I������
  '-------------------------------------------------------------------------
  i_dat = CnvZenAlphamericToHanEx(i_dat) '�S�p�𔼊p�ɕϊ�
  i_dat = Trimming(i_dat) '00FF��FF�̂悤�ɁA�擪��0������Ώ���
  
  If LogicalCheck(i_pls, i_int, i_rdx, i_dat) = False Then
    Exit Sub '�I��
  End If
  
  '-------------------------------------------------------------------------
  '��ϊ�
  '-------------------------------------------------------------------------
  'A���@
  o_radix__2 = ��ϊ�_A.RDX_CHANGE_A(i_rdx, i_dat, "2�i��")  ' 2�i��������
  o_radix_10 = ��ϊ�_A.RDX_CHANGE_A(i_rdx, i_dat, "10�i��") '10�i��������
  o_radix_16 = ��ϊ�_A.RDX_CHANGE_A(i_rdx, i_dat, "16�i��") '16�i��������
  
  'B���@(���ؗp)
'  Dim radix__2 As String
'  Dim radix_10 As String
'  Dim radix_16 As String
'  radix__2 = ��ϊ�_B.RDX_CHANGE_B(i_rdx, i_dat, "2�i��")    ' 2�i��������
'  radix_10 = ��ϊ�_B.RDX_CHANGE_B(i_rdx, i_dat, "10�i��")   '10�i��������
'  radix_16 = ��ϊ�_B.RDX_CHANGE_B(i_rdx, i_dat, "16�i��")   '16�i��������
'  Call ����.COMPARE(o_radix__2, radix__2) '�l��r
'  Call ����.COMPARE(o_radix_10, radix_10) '�l��r
'  Call ����.COMPARE(o_radix_16, radix_16) '�l��r
  
  '-------------------------------------------------------------------------
  '�ϊ����ʂ�z��Ɋi�[�B�d��������΁A�o���p�x���J�E���g����
  '-------------------------------------------------------------------------
  Call INPUT_ARR(o_radix__2, o_radix_10, o_radix_16)
  
  '-------------------------------------------------------------------------
  '�z��̕��ёւ�(�o���p�x�~���\�[�g)
  '-------------------------------------------------------------------------
  Call SORT_ARR
  
  '-------------------------------------------------------------------------
  '��������H
  '  �\���p�ɕ�������H�B���I�z��Ɍ��ʂ��i�[�����W���[���I������
  '-------------------------------------------------------------------------
  ARR_STR_DISP = ARR_STR   '�l�n���ŃR�s�[
  Call SplitStr(0, 4, "_") ' 2�i���������4������؂��"_"��}��
  Call SplitStr(1, 3, ",") '10�i���������3������؂��","��}��
  Call SplitStr(2, 4, "_") '16�i���������4������؂��"_"��}��
  
  '-------------------------------------------------------------------------
  '���C���V�[�g�ւ̏�������
  '-------------------------------------------------------------------------
  '�O��̕ϊ����ʂ��N���A
  Call RESULT_CLR
  
  '���͊���킩��悤�����F��ύX
  Call ResultColor(i_rdx, vbRed) 'vbRed/vbBlue/vbGreen �����R��
  
  '�ϊ���̕�����𕪉����Ȃ���1�������\��
  Call StringDecomposition(StrReverse(o_radix__2), Pg_Result_SttRng.Offset(0, 1))
  Call StringDecomposition(StrReverse(o_radix_10), Pg_Result_SttRng.Offset(1, 1))
  Call StringDecomposition(StrReverse(o_radix_16), Pg_Result_SttRng.Offset(2, 1))

  '�����L���O�X�V
  Call RANKING_WRITE(Pg_Ranking_Main_Stt)
    
  '-------------------------------------------------------------------------
  '�f�[�^�x�[�X�V�[�g�ւ̗�����������
  '-------------------------------------------------------------------------
  Call DATABASE_WRITE(o_radix__2, o_radix_10, o_radix_16)
  
End Sub



'###########################################################################
'
'�ȉ��A�֐�
'
'###########################################################################
'---------------------------------------------------------------------------
'���p�֐��B�S�p�����p�ϊ�
'  �Ή��\���画��
'  ����1   (IN)    �F�S�p���܂ޕ�����
'  �߂�l  (OUT)   �F���p������
'---------------------------------------------------------------------------
Private Function CnvZenAlphamericToHanEx(ByVal a_sZen As String) As String
    Dim sZenList As String '�S�p������
    Dim sHanList As String '���p������
    Dim sZenAr() As String '�S�p�����z��
    Dim sHanAr() As String '���p�����z��
    Dim sZen     As String '�S�p����
    Dim sHan     As String '���p����
    Dim iLen     As Long   '������
    Dim i        As Long
    
    '�Ή����X�g
    sZenList = "�`�a�b�c�d�e�������������O�P�Q�R�S�T�U�V�W�X�D�|"
    sHanList = "ABCDEFabcdef0123456789.-"
  
    '���������擾
    iLen = Len(sZenList)
    
    '�z�񒷃��T�C�Y
    ReDim sZenAr(iLen)
    ReDim sHanAr(iLen)
    
    For i = 0 To iLen - 1
        '���X�g(�S�p)��z��Ɋi�[
        sZenAr(i) = Mid(sZenList, i + 1, 1)
        '���X�g(���p)��z��Ɋi�[
        sHanAr(i) = Mid(sHanList, i + 1, 1)
    Next i
        
    '���͕�������Z�b�g
    CnvZenAlphamericToHanEx = a_sZen
    
    For i = 0 To iLen - 1
        '�S�p������Δ��p�ɒu��
        CnvZenAlphamericToHanEx = Replace(CnvZenAlphamericToHanEx, _
                                          sZenAr(i), sHanAr(i))
    Next i
End Function

'---------------------------------------------------------------------------
'�g���~���O  00FF��FF,�擪��0����������
'  ����1   (IN)    �F������
'  �߂�l  (OUT)   �F������
'---------------------------------------------------------------------------
Private Function Trimming(ByVal i_dat As String)
  Dim i As Long
  
  For i = 1 To Len(i_dat)
    If (Left(i_dat, 1) <> 0) Then
      Exit For
    Else
      i_dat = Mid(i_dat, 2, Len(i_dat) - 1)
    End If
  Next i
  
  If i_dat = "" Then
    i_dat = "0" '0000000 �� 0 �Ƃ���
  End If
  
  Trimming = i_dat
End Function


'---------------------------------------------------------------------------
'���͕������_���`�F�b�N
'
'  ����1   (IN)    �F����(�}) Tlue:��,   False:�����܂�
'  ����2   (IN)    �F�����L�� Tlue:����, False:�����܂�
'  ����3   (IN)    �F�ϊ��O�̊ "2�i��" or "10�i��" or "16�i��"
'  ����4   (IN)    �F���͕�����
'  �߂�l  (OUT)   �FTlue:�}�N�����s, False:�}�N���I��
'---------------------------------------------------------------------------
Private Function LogicalCheck(ByVal i_pls As Boolean, _
                              ByVal i_int As Boolean, _
                              ByVal i_rdx As String, _
                              ByVal i_dat As String) As Boolean
  Dim i As Long
  Dim cher As String '1�����B���͕������cher�Ɋi�[���āA1�������`�F�b�N�B
  LogicalCheck = True 'True�̂܂܊֐����I������Ύ��s�ł���B
  
  If i_dat = "" Then '�l�̂Ȃ����͂��͂���
    MsgBox "�l���󗓂ł��B"
    LogicalCheck = False
    Exit Function
  End If

  Select Case i_rdx
    Case "2�i��"
    
      If Len(i_dat) > 22 Then '23�����ȏ�̓��͂��͂���
      MsgBox "22�����ɂ����߂Ă�������"
      LogicalCheck = False
      Exit Function
      End If
    
      For i = 1 To Len(i_dat)
        cher = Mid(i_dat, i, 1)
        If cher <> "0" And cher <> "1" Then '0��1�ȊO�̓��͂��͂���
          MsgBox "0��1�̂ݓ��͂�������"
          LogicalCheck = False
          Exit Function
        End If
      Next i
      
      
    Case "10�i��"
    
      If Len(i_dat) > 7 Then '8�����ȏ�̓��͂��͂���
        MsgBox "4194303�܂œ��͉\�ł��B"
        LogicalCheck = False
        Exit Function
      End If
      
      If IsNumeric(i_dat) = False Then '������𐔒l�Ƃ��ĕ]���ł��Ȃ����͂��͂���
        MsgBox "���l����͂�������"
        LogicalCheck = False
        Exit Function
      End If
      
      If CLng(i_dat) > 4194303 Then '22bit�̍ő�l����������͂��͂���
        MsgBox "4194303�܂œ��͉\�ł��B"
        LogicalCheck = False
        Exit Function
      End If
      
    Case "16�i��"
    
      If (Len(i_dat) > 6) Then '7�����ȏ�̓��͂��͂���
        MsgBox "3fffff�܂œ��͉\�ł�"
        LogicalCheck = False
        Exit Function
      End If
      
      For i = 1 To Len(i_dat)
        cher = Mid(i_dat, i, 1)
        If cher Like "[!0-9a-fA-F]" Then '0�`9,a�`f�ȊO�̓��͂��͂���
          MsgBox "0�`9, a�`f �œ��͂�������"
          LogicalCheck = False
          Exit Function
        End If
      Next i
      
      '6�����̎��A1�����ڂ�0�`3�ȊO�̓I�[�o�[�t���[�Ƃ���
      If (Len(i_dat) = 6) And (Mid(i_dat, 1, 1) Like "[!0-3]") Then
        MsgBox "3fffff�܂œ��͉\�ł�"
        LogicalCheck = False
        Exit Function
      End If

    Case Else
      MsgBox "����v���_�E������I����������"
      LogicalCheck = False
      
    End Select
  
End Function


'---------------------------------------------------------------------------
'�ϊ���̒l��z��Ɋi�[, �o���p�x���J�E���g
'  ����1   (IN)    �F�ϊ���� 2�i��������
'  ����2   (IN)    �F�ϊ����10�i��������
'  ����3   (IN)    �F�ϊ����16�i��������
'  �߂�l  (OUT)   �F�Ȃ�
'---------------------------------------------------------------------------
Private Sub INPUT_ARR(ByVal str_rdx_2 As String, _
                      ByVal str_rdx10 As String, _
                      ByVal str_rdx16 As String)
  Dim i As Long
  
  '�T��..�󂢂Ă�z��ɕ�������i�[
  For i = LBound(ARR_STR, 1) To UBound(ARR_STR, 1)
    If ARR_STR(i, 1) = "" Then  '�l���
      ARR_STR(i, 0) = str_rdx_2
      ARR_STR(i, 1) = str_rdx10
      ARR_STR(i, 2) = str_rdx16
    End If
    
    '�d������������J�E���g�A�b�v���ă��[�v�𔲂���
    If str_rdx10 = ARR_STR(i, 1) Then
      ARR_CNT(i) = ARR_CNT(i) + 1 'COUNT UP
      Exit For
    End If
  Next i
  
End Sub


'---------------------------------------------------------------------------
'�z��̕��ёւ�
'  ���͕�����̏o���p�x���A�~���Ń\�[�g����
'��(�z��-1)!���������Ă��܂��B14!=800���B�����������Ȃ�悤���P�B2022_0120
'  ����    (IN)    �F�Ȃ�
'  �߂�l  (OUT)   �F�Ȃ�
'---------------------------------------------------------------------------
Sub SORT_ARR()

  Dim vSwap_str(2) As String '0:2�i��, 1:10�i��, 2:16�i��
  Dim vSwap_cnt As Long
  Dim i, j, k As Long
  
  For i = LBound(ARR_STR, 1) To UBound(ARR_STR, 1)
  
    If (ARR_CNT(i) = 0) Then
      Exit For '�T���I��
    End If
    
    If (ARR_CNT(i) < ARR_CNT(i + 1)) Then '�\�[�g�J�n
      For j = i + 1 To 1 Step -1
        If (ARR_CNT(j) > ARR_CNT(j - 1)) Then 'SWAP
          For k = 0 To 2
            'ARR_STR(k)
            vSwap_str(k) = ARR_STR(j - 1, k)
            ARR_STR(j - 1, k) = ARR_STR(j, k)
            ARR_STR(j, k) = vSwap_str(k)
          Next k
  
          'ARR_CNT
          vSwap_cnt = ARR_CNT(j - 1)
          ARR_CNT(j - 1) = ARR_CNT(j)
          ARR_CNT(j) = vSwap_cnt
        End If
      Next j
      Exit For '�\�[�g����
    End If
  Next i

'  '�o�u���\�[�g  �����������
'  For i = UBound(ARR_STR, 1) To LBound(ARR_STR, 1) Step -1 '�T��
'    For j = 0 To (i - 1)
'      '�召�֌W�s��v�ŕ��ёւ������{
'      If (ARR_CNT(j) < ARR_CNT(j + 1)) And (ARR_CNT(j) <> 0) Then
'        For k = 0 To 2
'
'        'ARR_STR(k)
'        vSwap_str(k) = ARR_STR(j, k)
'        ARR_STR(j, k) = ARR_STR(j + 1, k)
'        ARR_STR(j + 1, k) = vSwap_str(k)
'        Next k
'
'        'ARR_CNT
'        vSwap_cnt = ARR_CNT(j)
'        ARR_CNT(j) = ARR_CNT(j + 1)
'        ARR_CNT(j + 1) = vSwap_cnt
'
'      End If
'    Next j
'  Next i
    
End Sub


'---------------------------------------------------------------------------
'������ɁA�w�肵�������Ԋu�ŋ�؂蕶����}��
'
'  ����1   (IN)    �F�i���w��      0:2�i��, 1:10�i��, 2:16�i��
'  ����2   (IN)    �F��؂�Ԋu
'  ����3   (IN)    �F���ŋ�؂邩  "_"/","��
'  �߂�l  (OUT)   �F�Ȃ�
'---------------------------------------------------------------------------
Private Sub SplitStr(ByVal RdxNo As Integer, _
                     ByVal StrLength As Long, _
                     ByVal Char As String)
  
  Dim stt_mid As Long   'MID�֐��̃X�^�[�g�ʒu
  Dim o_str As String   '���ʊi�[�p
  Dim ModStrLen As Long '���܂�̕�����
  Dim i, j As Long
  
  For i = LBound(ARR_STR_DISP, 1) To UBound(ARR_STR_DISP, 1)
    If (ARR_CNT(i) = 0) Then
      Exit For '�I��
    Else
      '������
      stt_mid = 1
      o_str = ""
      ModStrLen = Len(ARR_STR_DISP(i, RdxNo)) Mod StrLength
        
      If (ModStrLen <> 0) Then
        stt_mid = stt_mid + ModStrLen
        o_str = Mid(ARR_STR_DISP(i, RdxNo), 1, ModStrLen)
      End If
        
      'xx�������̊Ԋu�ŁA"_"��","����������ł���
      For j = stt_mid To Len(ARR_STR_DISP(i, RdxNo)) Step StrLength
        If (j = 1) Then
          o_str = o_str & Mid(ARR_STR_DISP(i, RdxNo), j, StrLength)
        Else             '��"_"��","
          o_str = o_str & Char & Mid(ARR_STR_DISP(i, RdxNo), j, StrLength)
        End If
      Next j
        
      ARR_STR_DISP(i, RdxNo) = o_str
    End If
  Next i
End Sub


'---------------------------------------------------------------------------
'�������1�������������Ȃ��珑������
'  ����1   (IN)    �F������
'  ����2   (IN)    �F�������݉ӏ�
'  �߂�l  (OUT)   �F�Ȃ�
'---------------------------------------------------------------------------
Private Sub StringDecomposition(ByVal str As String, ByVal rng As Range)
  Dim i As Long
  
  With rng
    For i = 1 To Len(str) '������̒��������[�v
      .Offset(0, -i).Value = Mid(str, i, 1) '�ꕶ������������
   Next i
  End With

End Sub


'---------------------------------------------------------------------------
'�����F����
'  ����1   (IN)    �F�ϊ��O�̊
'  ����2   (IN)    �F�ύX�����������F
'  �߂�l  (OUT)   �F�Ȃ�
'---------------------------------------------------------------------------
Sub ResultColor(ByVal i_rdx As String, ByVal i_clolor As Long)
  Select Case i_rdx
    Case "2�i��"
      Pg_Result_Range().Rows(1).Font.Color = i_clolor ' 2�i��
  
    Case "10�i��"
      Pg_Result_Range().Rows(2).Font.Color = i_clolor '10�i��
    
    Case "16�i��"
      Pg_Result_Range().Rows(3).Font.Color = i_clolor '16�i��
  
  End Select
End Sub


'---------------------------------------------------------------------------
'�����񐔃����L���O��������
'  ����1   (IN)    �F�������݊J�n�̃Z���ʒu
'  �߂�l  (OUT)   �F�Ȃ�
'---------------------------------------------------------------------------
Private Sub RANKING_WRITE(ByVal stt_rng As Range)
  Dim i, j As Long
  
    '�������݊J�n�ʒu(stt_rng)����ɁA�����p�x�̃����L���O���������ށB
    For i = LBound(ARR_STR_DISP, 1) To UBound(ARR_STR_DISP, 1)   '�s����
      
      If (ARR_CNT(i) = 0) Or (i >= RANK_DISP_NUM_MAX) Then
        Exit For '�I��
      Else
        For j = LBound(ARR_STR_DISP, 2) To UBound(ARR_STR_DISP, 2) '�����
          stt_rng.Offset(i, j).Value = ARR_STR_DISP(i, j)
        Next j
          stt_rng.Offset(i, 3).Value = ARR_CNT(i) & "��"
      End If
    Next i
    
    '�����񒷂ɍ��킹�āA�񕝂���������
    For i = 2 To 5
      stt_rng.CurrentRegion.Columns(i).AutoFit
    Next i
    
End Sub


'---------------------------------------------------------------------------
'�f�[�^�x�[�X�ɗ�������������
'  ����1   (IN)    �F�ϊ���� 2�i��������
'  ����2   (IN)    �F�ϊ����10�i��������
'  ����3   (IN)    �F�ϊ����16�i��������
'  �߂�l  (OUT)   �F�Ȃ�
'---------------------------------------------------------------------------
Private Sub DATABASE_WRITE(ByVal o_radix__2 As String, _
                           ByVal o_radix_10 As String, _
                           ByVal o_radix_16 As String)
  
  Dim w_rw As Long '�������ލs(����)
  Dim i As Long
  
  '�ŏI�s��1�s���̍s�����擾
  w_rw = Pg_WSobj_DB.Cells(Rows.count, "B").End(xlUp).Row + 1
  
  Pg_WSobj_DB.Cells(w_rw, "B").Value = o_radix__2 ' 2�i������
  Pg_WSobj_DB.Cells(w_rw, "C").Value = o_radix_10 '10�i������
  Pg_WSobj_DB.Cells(w_rw, "D").Value = o_radix_16 '16�i������
  Pg_WSobj_DB.Cells(w_rw, "E").Value = Date       '���s��
  
End Sub


'---------------------------------------------------------------------------
'���C���V�[�g : �ϊ����ʃZ�����N���A
'  ����1   (IN)    �F�Ȃ�
'  �߂�l  (OUT)   �F�Ȃ�
'---------------------------------------------------------------------------
Public Sub RESULT_CLR()
  
  '�ϊ��O�̒l���̓Z����I��
  If ActiveSheet.Name = Pg_WSName_Main Then
    With Pg_I_DAT
    .Select  '�l���̓Z���ɃJ�[�\�����킹��
    End With
  End If
  
  '���ʕ\���Z���̒l���N���A
  With Pg_Result_Range()
    .ClearContents '�����폜
    .Font.ColorIndex = xlAutomatic '�����F������(��)��
  End With
  
End Sub


'---------------------------------------------------------------------------
'�f�[�^�x�[�X�V�[�g�F�ϊ��������N���A
'  ����1   (IN)    �F�Ȃ�
'  �߂�l  (OUT)   �F�Ȃ�
'---------------------------------------------------------------------------
Public Sub DATABASE_CLR()
  Dim return_msg As VbMsgBoxResult  'VbMsgBoxResult�񋓑�
  Dim i As Long
  Dim j As Long
  
  return_msg = MsgBox("�폜���܂��B�X�����ł����H", _
                      vbYesNo, "�m�F")
  
  If return_msg = vbYes Then
    '�z�񏉊���
    For i = LBound(ARR_STR, 1) To UBound(ARR_STR, 1)   '�s����
      For j = LBound(ARR_STR, 2) To UBound(ARR_STR, 2) '�����
        ARR_STR(i, j) = ""
      Next j
        ARR_CNT(i) = 0
    Next i
    
    '�f�[�^�x�[�X�͈͂̒l���N���A
    Pg_History_DB_Stt.CurrentRegion.Offset(1, 0).ClearContents
    
    '�����L���O�\���͈͂��N���A
    Pg_Ranking_Main_Stt.CurrentRegion.Offset(1, 1).ClearContents '���C���V�[�g
    'Pg_Ranking_DB_Stt.CurrentRegion.Offset(1, 1).ClearContents   '�����V�[�g
    
    '�J�[�\���ʒu�̒���
    If ActiveSheet.Name = Pg_WSName_Main Then
      Pg_I_DAT.Select  '�l���̓Z���ɃJ�[�\�����킹��
    Else
      Range("A1").Select 'A1�Z���I��
    End If
  End If
  
End Sub
