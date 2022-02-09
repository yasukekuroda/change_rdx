Attribute VB_Name = "ユーザー設定"
Option Explicit

'###########################################################################
'ユーザー設定箇所
'
'レイアウトを変更した場合、以下で指定している表示セル位置を変更する。
'設定変更後、"履歴削除"ボタンを押すことで設定を反映
'###########################################################################

'基数変換を行うシート名
Public Const Pg_WSName_Main As String = "基数変換"

'変換履歴を表示するシート名
Public Const Pg_WSName_DB As String = "使い方"

'最大何種類のデータを保存するか
Public Const ARR_MAX = 10000

'ランキングを何位まで表示するか
Public Const RANK_DISP_NUM_MAX = 10
'---------------------------------------------------------------------------
'
'メインシート:入力記入位置
'
'---------------------------------------------------------------------------
Public Function Pg_I_PLS() As Range
  Set Pg_I_PLS = Pg_WSobj_Main.Range("C1") '入力設定 : 符号(±) ///未実装
End Function
Public Function Pg_I_INT() As Range
  Set Pg_I_INT = Pg_WSobj_Main.Range("C2") '入力設定 : 小数     ///未実装
End Function
Public Function Pg_I_RDX() As Range
  Set Pg_I_RDX = Pg_WSobj_Main.Range("C4") '変換前の値 : 基数
End Function
Public Function Pg_I_DAT() As Range
  Set Pg_I_DAT = Pg_WSobj_Main.Range("C5") '変換前の値 : 値
End Function
'---------------------------------------------------------------------------
'
'メインシート:変換結果を1文字ずつ表示するエリア
'
'---------------------------------------------------------------------------
Public Function Pg_Result_Range() As Range
  Set Pg_Result_Range = Pg_WSobj_Main.Range("F4:AA6") 'F4セル～AA6セル
End Function
'---------------------------------------------------------------------------
'
'メインシート:ランキングの書き込み開始セル
'
'---------------------------------------------------------------------------
Public Function Pg_Ranking_Main_Stt() As Range
  Set Pg_Ranking_Main_Stt = Pg_WSobj_Main.Cells(5, "AD") 'AD5セル
End Function
'---------------------------------------------------------------------------
'
'データベースシート:変換履歴の書き込み開始セル
'
'---------------------------------------------------------------------------
Public Function Pg_History_DB_Stt() As Range
  Set Pg_History_DB_Stt = Pg_WSobj_DB.Range("B4") 'B4セル
End Function
'---------------------------------------------------------------------------
'
'データベースシート:ランキングの書き込みセル位置
'
'---------------------------------------------------------------------------
Public Function Pg_Ranking_DB_Stt() As Range
  Set Pg_Ranking_DB_Stt = Pg_WSobj_DB.Cells(5, "H") 'H5セル
End Function
'---------------------------------------------------------------------------
'
'ワークシートオブジェクト定義:メインシート
'
'---------------------------------------------------------------------------
Public Function Pg_WSobj_Main() As Worksheet
  Set Pg_WSobj_Main = Worksheets(Pg_WSName_Main)
End Function
'---------------------------------------------------------------------------
'
'ワークシートオブジェクト定義:データベースシート
'
'---------------------------------------------------------------------------
Public Function Pg_WSobj_DB() As Worksheet
  Set Pg_WSobj_DB = Worksheets(Pg_WSName_DB)
End Function
'---------------------------------------------------------------------------
'
'メインシート:選択エリアの右上セルを取得
'
'---------------------------------------------------------------------------
Public Function Pg_Result_SttRng() As Range
  Set Pg_Result_SttRng = Pg_WSobj_Main.Cells(Pg_Result_Range().Rows(1).Row, _
    Pg_Result_Range().Columns(Pg_Result_Range().Columns.count).Column)
End Function
