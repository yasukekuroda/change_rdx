Attribute VB_Name = "���[�U�[�ݒ�"
Option Explicit

'###########################################################################
'���[�U�[�ݒ�ӏ�
'
'���C�A�E�g��ύX�����ꍇ�A�ȉ��Ŏw�肵�Ă���\���Z���ʒu��ύX����B
'�ݒ�ύX��A"�����폜"�{�^�����������ƂŐݒ�𔽉f
'###########################################################################

'��ϊ����s���V�[�g��
Public Const Pg_WSName_Main As String = "��ϊ�"

'�ϊ�������\������V�[�g��
Public Const Pg_WSName_DB As String = "�g����"

'�ő剽��ނ̃f�[�^��ۑ����邩
Public Const ARR_MAX = 10000

'�����L���O�����ʂ܂ŕ\�����邩
Public Const RANK_DISP_NUM_MAX = 10
'---------------------------------------------------------------------------
'
'���C���V�[�g:���͋L���ʒu
'
'---------------------------------------------------------------------------
Public Function Pg_I_PLS() As Range
  Set Pg_I_PLS = Pg_WSobj_Main.Range("C1") '���͐ݒ� : ����(�}) ///������
End Function
Public Function Pg_I_INT() As Range
  Set Pg_I_INT = Pg_WSobj_Main.Range("C2") '���͐ݒ� : ����     ///������
End Function
Public Function Pg_I_RDX() As Range
  Set Pg_I_RDX = Pg_WSobj_Main.Range("C4") '�ϊ��O�̒l : �
End Function
Public Function Pg_I_DAT() As Range
  Set Pg_I_DAT = Pg_WSobj_Main.Range("C5") '�ϊ��O�̒l : �l
End Function
'---------------------------------------------------------------------------
'
'���C���V�[�g:�ϊ����ʂ�1�������\������G���A
'
'---------------------------------------------------------------------------
Public Function Pg_Result_Range() As Range
  Set Pg_Result_Range = Pg_WSobj_Main.Range("F4:AA6") 'F4�Z���`AA6�Z��
End Function
'---------------------------------------------------------------------------
'
'���C���V�[�g:�����L���O�̏������݊J�n�Z��
'
'---------------------------------------------------------------------------
Public Function Pg_Ranking_Main_Stt() As Range
  Set Pg_Ranking_Main_Stt = Pg_WSobj_Main.Cells(5, "AD") 'AD5�Z��
End Function
'---------------------------------------------------------------------------
'
'�f�[�^�x�[�X�V�[�g:�ϊ������̏������݊J�n�Z��
'
'---------------------------------------------------------------------------
Public Function Pg_History_DB_Stt() As Range
  Set Pg_History_DB_Stt = Pg_WSobj_DB.Range("B4") 'B4�Z��
End Function
'---------------------------------------------------------------------------
'
'�f�[�^�x�[�X�V�[�g:�����L���O�̏������݃Z���ʒu
'
'---------------------------------------------------------------------------
Public Function Pg_Ranking_DB_Stt() As Range
  Set Pg_Ranking_DB_Stt = Pg_WSobj_DB.Cells(5, "H") 'H5�Z��
End Function
'---------------------------------------------------------------------------
'
'���[�N�V�[�g�I�u�W�F�N�g��`:���C���V�[�g
'
'---------------------------------------------------------------------------
Public Function Pg_WSobj_Main() As Worksheet
  Set Pg_WSobj_Main = Worksheets(Pg_WSName_Main)
End Function
'---------------------------------------------------------------------------
'
'���[�N�V�[�g�I�u�W�F�N�g��`:�f�[�^�x�[�X�V�[�g
'
'---------------------------------------------------------------------------
Public Function Pg_WSobj_DB() As Worksheet
  Set Pg_WSobj_DB = Worksheets(Pg_WSName_DB)
End Function
'---------------------------------------------------------------------------
'
'���C���V�[�g:�I���G���A�̉E��Z�����擾
'
'---------------------------------------------------------------------------
Public Function Pg_Result_SttRng() As Range
  Set Pg_Result_SttRng = Pg_WSobj_Main.Cells(Pg_Result_Range().Rows(1).Row, _
    Pg_Result_Range().Columns(Pg_Result_Range().Columns.count).Column)
End Function
