Attribute VB_Name = "��ϊ�_B"
Option Explicit

Private i As Long
Private varBinary As Variant
Private colHValue As New Collection '�A�z�z��ACollection�I�u�W�F�N�g�̍쐬
Private lngNu() As Long


'---------------------------------------------------------------------------
'��ϊ��Ǘ� ..B���@
'���p
'  �����񁩁�16�i������2�i���̑��ݕϊ�
'  https://excel.syogyoumujou.com/memorandum/hex_binary.html
'  ��i������\�i���𓾂郆�[�U��`�֐�
'  https://www.moug.net/tech/exvba/0100013.html
'
'  ����1   (IN)    �F�ϊ��O�̊
'  ����2   (IN)    �F���͒l
'  ����3   (IN)    �F�ϊ���̊ "2�i��" or "10�i��" or "16�i��"
'  �߂�l  (OUT)   �F�ϊ�����������
'---------------------------------------------------------------------------
'������16�i������2�i���̑��ݕϊ�
Public Function RDX_CHANGE_B(ByVal i_rdx As String, _
                             ByVal i_dat As String, _
                             ByVal o_rdx As String)
                             
  Dim strData As String
  
  '2����16 �i���̕ϊ������X�g�őΉ��Â���B
  '�ϊ����ʂŕs�v��0���g���~���O���A�o�́B
  varBinary = Array("0000", "0001", "0010", "0011", _
                    "0100", "0101", "0110", "0111", _
                    "1000", "1001", "1010", "1011", _
                    "1100", "1101", "1110", "1111")
                    
  Set colHValue = New Collection '������
  
  '�A�z�z��Ɂu�L�[�v�Ƃ���varBinary��2�i���A�u�A�C�e���v�Ƃ��đΉ�����16�i���u0�`F�v���i�[
  For i = 0 To 15
    colHValue.Add CStr(Hex$(i)), varBinary(i)
  Next
    
  Select Case i_rdx
    Case "2�i��"
      If o_rdx = "2�i��" Then
        RDX_CHANGE_B = i_dat                  '2 ��  2
      ElseIf o_rdx = "10�i��" Then
        RDX_CHANGE_B = SampleBinToDeci(i_dat) '2 �� 10
      Else
        RDX_CHANGE_B = BtoH(i_dat)            '2 �� 16
      End If
      
    Case "10�i��"
      If o_rdx = "2�i��" Then
        RDX_CHANGE_B = ExDeciToBin(CLng(i_dat)) '10 ��  2
      ElseIf o_rdx = "10�i��" Then
        RDX_CHANGE_B = i_dat                    '10 �� 10
      Else
        RDX_CHANGE_B = Hex(i_dat)               '10 �� 16
      End If
    
    Case "16�i��"
      If o_rdx = "2�i��" Then
        RDX_CHANGE_B = Trimming(HtoB(i_dat))    '16 ��  2
      ElseIf o_rdx = "10�i��" Then
        RDX_CHANGE_B = HexToDec(i_dat)          '16 �� 10
      Else
        RDX_CHANGE_B = UCase(i_dat)             '16 �� 16
      End If
      
  End Select
  
  Erase lngNu '�z��̉��...�Ȃ��Ă���������

End Function


'---------------------------------------------------------------------------
'��ϊ�(16��2)
'  ����1   (IN)    �F16�i��������
'  �߂�l  (OUT)   �F 2�i��������
'---------------------------------------------------------------------------
Private Function HtoB(ByVal strH As String) As String '16�i����2�i��
    ReDim strHtoB(1 To Len(strH)) As String
    For i = 1 To Len(strH)
        strHtoB(i) = varBinary(Val("&h" & Mid$(strH, i, 1)))
    Next
    HtoB = Join$(strHtoB, vbNullString)
End Function


'---------------------------------------------------------------------------
'��ϊ�(2��16)
'  ����1   (IN)    �F 2�i��������
'  �߂�l  (OUT)   �F16�i��������
'---------------------------------------------------------------------------
Private Function BtoH(ByVal strB As String) As String '2�i����16�i��
  '��������4�̔{���ɒ���
  If ((Len(strB) Mod 4) > 0) Then
    strB = Left("0000", 4 - (Len(strB) Mod 4)) & strB
  End If
  
  ReDim strBtoH(1 To Len(strB) / 4) As String
  For i = 1 To Len(strB) / 4 '2�i��(4bit��)��16�i���ɕϊ�
    strBtoH(i) = colHValue.Item(Mid$(strB, (i - 1) * 4 + 1, 4))
  Next
  BtoH = Join$(strBtoH, vbNullString)
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
    i_dat = "0"
  End If
  
  Trimming = i_dat
End Function

'�ő�̂Q�ׂ̂���̒l��T��
Private Function Ex2noBeki(deci As Long) As Integer
    Dim i As Integer
    
    i = 0
    Do
        'deci���傫��
        If deci < 2 ^ i Then
            '���̈�O�ׂ̂���
            Ex2noBeki = i - 1
            Exit Function
        End If
        i = i + 1
    Loop
End Function

'---------------------------------------------------------------------------
'��ϊ�(10��2)
'  ����1   (IN)    �F10�i�����l
'  �߂�l  (OUT)   �F 2�i��������
'---------------------------------------------------------------------------
Private Function ExDeciToBin(deci As Long) As String
    Dim ln As Long
    Dim stemp As String
    Dim i As Integer
    Dim count As Integer
    
    stemp = "1"
    'deci��菬�����A�ő�̂Q�ׂ̂���̒l��T��
    count = Ex2noBeki(deci)
    ln = deci - 2 ^ count
    '�M�Z�Ɠ����悤�ɌJ��Ԃ�
    For i = count - 1 To 0 Step -1
         If ln < 2 ^ i Then
            stemp = stemp & "0"
         Else
            stemp = stemp & "1"
            ln = ln - (2 ^ i)
         End If
    Next i
    ExDeciToBin = stemp
End Function
'---------------------------------------------------------------------------
'��ϊ�(2��10)
'  ����1   (IN)    �F 2�i�����l
'  �߂�l  (OUT)   �F10�i��������
'---------------------------------------------------------------------------
Private Function SampleBinToDeci(Binary As String) As Long

Dim myLen As Integer
Dim i As Integer

    myLen = Len(Binary)
    For i = 1 To myLen
        If Mid(Binary, i, 1) = "1" Then
            SampleBinToDeci = SampleBinToDeci + 2 ^ (myLen - i)
        End If
    Next

End Function


'---------------------------------------------------------------------------
'���̊֐����� ��ϊ�_A�V�[�g�Ɠ����ł��邪�A Clng()�֐��𗘗p���Ă���B
'��ϊ�(16��10)
'  ����1   (IN)    �F16�i��������
'  �߂�l  (OUT)   �F10�i��������
'---------------------------------------------------------------------------
Private Function HexToDec(ByVal a_sHex As String)
  Dim dDec As Double   '10�i���l
    
  '10�i���l�ɕϊ�
  dDec = CLng("&H" & a_sHex)
    
  HexToDec = CStr(dDec)
End Function
