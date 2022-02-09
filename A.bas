Attribute VB_Name = "��ϊ�_A"
Option Explicit
'---------------------------------------------------------------------------
'��ϊ��Ǘ� ..A���@
'  �Q�l�T�C�g
'  Excel��Ƃ�VBA�Ō�����
'  https://vbabeginner.net/convert-hextodec/
'
'
'  ����1   (IN)    �F�ϊ��O�̊
'  ����2   (IN)    �F���͒l
'  ����3   (IN)    �F�ϊ���̊ "2�i��" or "10�i��" or "16�i��"
'  �߂�l  (OUT)   �F�ϊ�����������
'---------------------------------------------------------------------------
Public Function RDX_CHANGE_A(ByVal i_rdx As String, _
                             ByVal i_dat As String, _
                             ByVal o_rdx As String)
  Select Case i_rdx
    Case "2�i��"
      If o_rdx = "2�i��" Then
        RDX_CHANGE_A = i_dat           '2 ��  2
      ElseIf o_rdx = "10�i��" Then
        RDX_CHANGE_A = BinToDec(i_dat) '2 �� 10
      Else
        RDX_CHANGE_A = BinToHex(i_dat) '2 �� 16
      End If
      
    Case "10�i��"
      If o_rdx = "2�i��" Then
        RDX_CHANGE_A = DecToBin(i_dat) '10 ��  2
      ElseIf o_rdx = "10�i��" Then
        RDX_CHANGE_A = i_dat           '10 �� 10
      Else
        RDX_CHANGE_A = Hex(i_dat)      '10 �� 16
      End If
    
    Case "16�i��"
      If o_rdx = "2�i��" Then
        RDX_CHANGE_A = DecToBin(CStr(HexToDec(i_dat))) ' 16 �� 10 �� 2
      ElseIf o_rdx = "10�i��" Then
        RDX_CHANGE_A = HexToDec(i_dat)                  '16 �� 10
      Else
        RDX_CHANGE_A = UCase(i_dat)                     '16 �� 16
      End If
      
  End Select
End Function


'---------------------------------------------------------------------------
'��ϊ�(2��10)
'  ����1   (IN)    �F2�i��������
'  �߂�l  (OUT)   �F10�i��������
'---------------------------------------------------------------------------
Private Function BinToDec(ByVal I_BIN As String)
  Dim i        As Long
  Dim i_Len    As Long    '  2�i��������
  Dim sParts   As String   '  2�i���������؂�o�����ꕔ
  Dim O_DEC    As Long    ' 10�i���l
    
  O_DEC = 0 '������

  i_Len = Len(I_BIN) '���͂̕����񒷂��擾
    
  For i = 1 To i_Len
    '2�i���������1�����؂�o��
    sParts = Mid(I_BIN, i, 1)
        
    '2�i���l�~2��n��̒l�����Z����
    O_DEC = O_DEC + 2 ^ (i_Len - i) * CLng(sParts)
  Next i
    
  '������Ƃ��ďo��
  BinToDec = CStr(O_DEC)
End Function


'---------------------------------------------------------------------------
'��ϊ�(2��16)
'  ����1   (IN)    �F2�i��������
'  �߂�l  (OUT)   �F16�i��������
'---------------------------------------------------------------------------
Private Function BinToHex(ByVal I_BIN As String)
  Dim i           As Long
  Dim iParts      As Long    '  2�i�������񃋁[�v�J�E���^
  Dim sParts      As String  '  2�i���������؂�o����4����
  Dim iRemainder  As Long    ' �]��
  Dim iDec        As Long    ' 10�i���l
    
  '2�i��������̕�������4�Ŋ������]����擾
  iRemainder = Len(I_BIN) Mod 4
    
  '�]�肪����ꍇ�A�s�����Ă���"0"��t�^
  If (iRemainder > 0) Then
    I_BIN = Left("0000", 4 - iRemainder) & I_BIN
  End If
    
  '2�i���������������4���������[�v
  For iParts = 1 To Len(I_BIN) Step 4
    '2�i���������4�����؂�o��
    sParts = Mid(I_BIN, iParts, 4)
        
    '10�i���l��������
    iDec = 0
        
    '�؂�o����������������珇��1��������10�i���l�ɕϊ�����4�����������v����
    For i = 0 To 3
      '2�i���l�~2��n��̒l�����Z����
      iDec = iDec + 2 ^ (3 - i) * CInt(Mid(sParts, i + 1, 1))
    Next i
        
    '0000�ȊO�̓��͂ɑ΂��Ď��s
    If (BinToHex <> "" Or iDec <> 0) Then
      '16�i���������A��
      BinToHex = BinToHex & CStr(Hex(iDec))
    End If
  Next iParts
    
  '1'b0�̓��͂ɑ΂��Ă�0���o��
  If (BinToHex = "") Then
      BinToHex = "0"
  End If

End Function


'---------------------------------------------------------------------------
'��ϊ�(10��2)
'  ����1   (IN)    �F10�i��������
'  �߂�l  (OUT)   �F 2�i��������
'---------------------------------------------------------------------------
Public Function DecToBin(ByVal a_sDec As String)
  Dim i           As Long
  Dim iRemainder  As Long    '�]��
  Dim dDiv        As Double  '��
    
  '����10�i���������10�i���l�Ƃ��Ď擾
  dDiv = Val(a_sDec)
    
  '���������܂Ń��[�v
  Do
    '10�i���l��2�Ŋ������]����擾
    iRemainder = dDiv Mod 2
    
    '10�i���l��2�Ŋ����������擾�i�����[�v��10�i���l�ɂȂ�j
    dDiv = Int(dDiv / 2)
        
    '2�i��������̍��ɗ]���A��
    DecToBin = CStr(iRemainder) & DecToBin
        
    '10�i���l��2�����i����2�Ŋ���Ȃ��̂ł����Ń��[�v�I���j
    If (dDiv < 2) Then
      If (dDiv = 1) Then
        DecToBin = CStr(dDiv) & DecToBin '�ŏ�ʌ��̒l�Ƃ���"1"��A��
      End If
      Exit Do '���[�v�𔲂���
    End If
  Loop

End Function


'---------------------------------------------------------------------------
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
