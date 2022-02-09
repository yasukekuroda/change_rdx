Attribute VB_Name = "検証"
Option Explicit
'---------------------------------------------------------------------------
'検証
'  変換が正しくできているか期待値とあてて比較する。
'
'入力範囲                                     結果
'   2進数 : 0～11_1111_1111_1111_1111_1111  ||  OK
'  10進数 : 0～4,194,303                    ||  OK
'  16進数 : 0～3F_FFFF                      ||  OK
'
'実行時間
'  1回目  31分20秒
'  2回目  24分40秒
'  3回目  19分40秒
'---------------------------------------------------------------------------
Public Sub INSPECTION()
  Dim i As Long
  Call マクロ高速化開始

  For i = 0 To 4194303 '正の整数の全入力パターン
  
    ' 2進数→10進数, 16進数
    '(A方法で変換した結果が引数にあるが、Main関数内で再度ABに変換し比較する)
    Call Main処理部.Main(True, True, "2進数", 基数変換_A.DecToBin(str(i)))
    
    '10進数→ 2進数, 16進数
    Call Main処理部.Main(True, True, "10進数", i)
    
    '16進数→ 2進数, 10進数
    Call Main処理部.Main(True, True, "16進数", Hex(i_dat))
  Next i
    
  Call マクロ高速化終了
  MsgBox "実行成功"
End Sub


'---------------------------------------------------------------------------
'値比較
'  引数1   (IN)    ：A方法の結果
'  引数2   (IN)    ：B方法の結果
'  戻り値  (OUT)   ：なし
'---------------------------------------------------------------------------
Public Sub COMPARE(ByVal str1 As String, _
                   ByVal str2 As String)
  If str1 <> str2 Then
    MsgBox "値不一致です。" & vbCrLf & _
           "str1 : " & str1 & vbCrLf & _
           "str2 : " & str2
  End If
End Sub
                
                
'##########################################################
'マクロ高速化と警告停止等
'##########################################################
Private Sub マクロ高速化開始()

Application.Interactive = False    'キーボードの入力OFF
Application.ScreenUpdating = False '画面描写を停止
Application.Cursor = xlWait        'カーソルを砂時計型に
Application.EnableEvents = False   'イベントを抑止
Application.DisplayAlerts = False  '警告メッセージ非表示
Application.Calculation = xlCalculationManual '計算を手動化

End Sub


Private Sub マクロ高速化終了()

'Application.StatusBar = False    'ステータスバー非表示
Application.Calculation = xlCalculationAutomatic '計算を自動
Application.DisplayAlerts = True  'メッセージ表示開始
Application.EnableEvents = True   'イベントを受付開始
Application.Cursor = xlDefault    'カーソルをデフォルト
Application.ScreenUpdating = True '画面描画開始
Application.Interactive = True    'キーボード入力受付開始

End Sub


'  'デバック : 配列の値表示
'  Dim i As Long
'  For i = LBound(ARR_STR, 1) To UBound(ARR_STR, 1)
'    If ARR_STR(i, 0) <> "" Then
'      MsgBox "ARR_STR[" & i & "][0] = """ & ARR_STR(i, 0) & """   " & _
'             "ARR_STR[" & i & "][1] = """ & ARR_STR(i, 1) & """   " & _
'             "ARR_STR[" & i & "][2] = """ & ARR_STR(i, 2) & """   " & _
'             "ARR_CNT[" & i & "] = " & ARR_CNT(i) & "回"
'    End If
'  Next i





