Attribute VB_Name = "Main処理部"
Option Explicit
'###########################################################################
'基数変換スクリプト
'2022/01/31   黒田
'符号の有無 / 小数点の有無  に未対応です。。
'
'引用
'  VBAで文字列の全角英数字を半角に変換する
'  https://vbabeginner.net/convert-full-width-alphanumeric-characters-in-a-character-string-to-half-width/
'###########################################################################
'---------------------------------------------------------------------------
'配列定義  (本モジュール内のみ有効)
'---------------------------------------------------------------------------
'変換後の文字列を格納 (xx,0)→2進数, (xx,1)→:10進数, (xx,2)→:16進数
Dim ARR_STR(ARR_MAX - 1, 2) As String

'画面描写用。動的。("_"や","を挿入)
Dim ARR_STR_DISP() As String

'出現回数をカウント。
'2進数で"111"と検索したものと10進数で"7"と検索したものは同一として扱う。
Dim ARR_CNT(ARR_MAX - 1) As Long


'###########################################################################
'
'Main 関数
'
'検証用
'Public Sub Main(ByVal i_pls As Boolean, _
'                ByVal i_int As Boolean, _
'                ByVal i_rdx As String, _
'                ByVal i_dat As String)
'###########################################################################
Public Sub Main()

  '-------------------------------------------------------------------------
  '変数宣言
  '-------------------------------------------------------------------------
  Dim i_pls As Boolean     'Tlue:正の数のみ,  False:負の数を含む
  Dim i_int As Boolean     'Tlue:整数,        False:小数を含む
  Dim i_rdx As String      '変換前の基数("2進数/10進数/16進数")
  Dim i_dat As String      '入力文字列
  '設定値読み込み
  i_pls = Pg_I_PLS.Value = "なし"
  i_int = Pg_I_INT.Value = "なし"
  i_rdx = Pg_I_RDX.Value
  i_dat = Pg_I_DAT.Value

  '変換後の文字列格納
  Dim o_radix__2 As String ' 2進数
  Dim o_radix_10 As String '10進数
  Dim o_radix_16 As String '16進数

  '-------------------------------------------------------------------------
  '入力値の論理チェック
  '  戻り値が False のとき、マクロを終了する
  '-------------------------------------------------------------------------
  i_dat = CnvZenAlphamericToHanEx(i_dat) '全角を半角に変換
  i_dat = Trimming(i_dat) '00FF→FFのように、先頭に0があれば除去
  
  If LogicalCheck(i_pls, i_int, i_rdx, i_dat) = False Then
    Exit Sub '終了
  End If
  
  '-------------------------------------------------------------------------
  '基数変換
  '-------------------------------------------------------------------------
  'A方法
  o_radix__2 = 基数変換_A.RDX_CHANGE_A(i_rdx, i_dat, "2進数")  ' 2進数←入力
  o_radix_10 = 基数変換_A.RDX_CHANGE_A(i_rdx, i_dat, "10進数") '10進数←入力
  o_radix_16 = 基数変換_A.RDX_CHANGE_A(i_rdx, i_dat, "16進数") '16進数←入力
  
  'B方法(検証用)
'  Dim radix__2 As String
'  Dim radix_10 As String
'  Dim radix_16 As String
'  radix__2 = 基数変換_B.RDX_CHANGE_B(i_rdx, i_dat, "2進数")    ' 2進数←入力
'  radix_10 = 基数変換_B.RDX_CHANGE_B(i_rdx, i_dat, "10進数")   '10進数←入力
'  radix_16 = 基数変換_B.RDX_CHANGE_B(i_rdx, i_dat, "16進数")   '16進数←入力
'  Call 検証.COMPARE(o_radix__2, radix__2) '値比較
'  Call 検証.COMPARE(o_radix_10, radix_10) '値比較
'  Call 検証.COMPARE(o_radix_16, radix_16) '値比較
  
  '-------------------------------------------------------------------------
  '変換結果を配列に格納。重複があれば、出現頻度をカウントする
  '-------------------------------------------------------------------------
  Call INPUT_ARR(o_radix__2, o_radix_10, o_radix_16)
  
  '-------------------------------------------------------------------------
  '配列の並び替え(出現頻度降順ソート)
  '-------------------------------------------------------------------------
  Call SORT_ARR
  
  '-------------------------------------------------------------------------
  '文字列加工
  '  表示用に文字列加工。動的配列に結果を格納→モジュール終了後解放
  '-------------------------------------------------------------------------
  ARR_STR_DISP = ARR_STR   '値渡しでコピー
  Call SplitStr(0, 4, "_") ' 2進数文字列に4文字区切りで"_"を挿入
  Call SplitStr(1, 3, ",") '10進数文字列に3文字区切りで","を挿入
  Call SplitStr(2, 4, "_") '16進数文字列に4文字区切りで"_"を挿入
  
  '-------------------------------------------------------------------------
  'メインシートへの書き込み
  '-------------------------------------------------------------------------
  '前回の変換結果をクリア
  Call RESULT_CLR
  
  '入力基数がわかるよう文字色を変更
  Call ResultColor(i_rdx, vbRed) 'vbRed/vbBlue/vbGreen 等自由に
  
  '変換後の文字列を分解しながら1文字ずつ表示
  Call StringDecomposition(StrReverse(o_radix__2), Pg_Result_SttRng.Offset(0, 1))
  Call StringDecomposition(StrReverse(o_radix_10), Pg_Result_SttRng.Offset(1, 1))
  Call StringDecomposition(StrReverse(o_radix_16), Pg_Result_SttRng.Offset(2, 1))

  'ランキング更新
  Call RANKING_WRITE(Pg_Ranking_Main_Stt)
    
  '-------------------------------------------------------------------------
  'データベースシートへの履歴書き込み
  '-------------------------------------------------------------------------
  Call DATABASE_WRITE(o_radix__2, o_radix_10, o_radix_16)
  
End Sub



'###########################################################################
'
'以下、関数
'
'###########################################################################
'---------------------------------------------------------------------------
'引用関数。全角→半角変換
'  対応表から判定
'  引数1   (IN)    ：全角を含む文字列
'  戻り値  (OUT)   ：半角文字列
'---------------------------------------------------------------------------
Private Function CnvZenAlphamericToHanEx(ByVal a_sZen As String) As String
    Dim sZenList As String '全角文字列挙
    Dim sHanList As String '半角文字列挙
    Dim sZenAr() As String '全角文字配列
    Dim sHanAr() As String '半角文字配列
    Dim sZen     As String '全角文字
    Dim sHan     As String '半角文字
    Dim iLen     As Long   '文字数
    Dim i        As Long
    
    '対応リスト
    sZenList = "ＡＢＣＤＥＦａｂｃｄｅｆ０１２３４５６７８９．－"
    sHanList = "ABCDEFabcdef0123456789.-"
  
    '文字数を取得
    iLen = Len(sZenList)
    
    '配列長リサイズ
    ReDim sZenAr(iLen)
    ReDim sHanAr(iLen)
    
    For i = 0 To iLen - 1
        'リスト(全角)を配列に格納
        sZenAr(i) = Mid(sZenList, i + 1, 1)
        'リスト(半角)を配列に格納
        sHanAr(i) = Mid(sHanList, i + 1, 1)
    Next i
        
    '入力文字列をセット
    CnvZenAlphamericToHanEx = a_sZen
    
    For i = 0 To iLen - 1
        '全角があれば半角に置換
        CnvZenAlphamericToHanEx = Replace(CnvZenAlphamericToHanEx, _
                                          sZenAr(i), sHanAr(i))
    Next i
End Function

'---------------------------------------------------------------------------
'トリミング  00FF→FF,先頭の0を除去する
'  引数1   (IN)    ：文字列
'  戻り値  (OUT)   ：文字列
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
    i_dat = "0" '0000000 → 0 とする
  End If
  
  Trimming = i_dat
End Function


'---------------------------------------------------------------------------
'入力文字列を論理チェック
'
'  引数1   (IN)    ：符号(±) Tlue:正,   False:負を含む
'  引数2   (IN)    ：小数有無 Tlue:整数, False:小数含む
'  引数3   (IN)    ：変換前の基数 "2進数" or "10進数" or "16進数"
'  引数4   (IN)    ：入力文字列
'  戻り値  (OUT)   ：Tlue:マクロ実行, False:マクロ終了
'---------------------------------------------------------------------------
Private Function LogicalCheck(ByVal i_pls As Boolean, _
                              ByVal i_int As Boolean, _
                              ByVal i_rdx As String, _
                              ByVal i_dat As String) As Boolean
  Dim i As Long
  Dim cher As String '1文字。入力文字列をcherに格納して、1文字ずつチェック。
  LogicalCheck = True 'Trueのまま関数を終えられれば実行できる。
  
  If i_dat = "" Then '値のない入力をはじく
    MsgBox "値が空欄です。"
    LogicalCheck = False
    Exit Function
  End If

  Select Case i_rdx
    Case "2進数"
    
      If Len(i_dat) > 22 Then '23文字以上の入力をはじく
      MsgBox "22文字におさめてください"
      LogicalCheck = False
      Exit Function
      End If
    
      For i = 1 To Len(i_dat)
        cher = Mid(i_dat, i, 1)
        If cher <> "0" And cher <> "1" Then '0と1以外の入力をはじく
          MsgBox "0と1のみ入力ください"
          LogicalCheck = False
          Exit Function
        End If
      Next i
      
      
    Case "10進数"
    
      If Len(i_dat) > 7 Then '8文字以上の入力をはじく
        MsgBox "4194303まで入力可能です。"
        LogicalCheck = False
        Exit Function
      End If
      
      If IsNumeric(i_dat) = False Then '文字列を数値として評価できない入力をはじく
        MsgBox "数値を入力ください"
        LogicalCheck = False
        Exit Function
      End If
      
      If CLng(i_dat) > 4194303 Then '22bitの最大値をこえる入力をはじく
        MsgBox "4194303まで入力可能です。"
        LogicalCheck = False
        Exit Function
      End If
      
    Case "16進数"
    
      If (Len(i_dat) > 6) Then '7文字以上の入力をはじく
        MsgBox "3fffffまで入力可能です"
        LogicalCheck = False
        Exit Function
      End If
      
      For i = 1 To Len(i_dat)
        cher = Mid(i_dat, i, 1)
        If cher Like "[!0-9a-fA-F]" Then '0～9,a～f以外の入力をはじく
          MsgBox "0～9, a～f で入力ください"
          LogicalCheck = False
          Exit Function
        End If
      Next i
      
      '6文字の時、1文字目が0～3以外はオーバーフローとする
      If (Len(i_dat) = 6) And (Mid(i_dat, 1, 1) Like "[!0-3]") Then
        MsgBox "3fffffまで入力可能です"
        LogicalCheck = False
        Exit Function
      End If

    Case Else
      MsgBox "基数をプルダウンから選択ください"
      LogicalCheck = False
      
    End Select
  
End Function


'---------------------------------------------------------------------------
'変換後の値を配列に格納, 出現頻度をカウント
'  引数1   (IN)    ：変換後の 2進数文字列
'  引数2   (IN)    ：変換後の10進数文字列
'  引数3   (IN)    ：変換後の16進数文字列
'  戻り値  (OUT)   ：なし
'---------------------------------------------------------------------------
Private Sub INPUT_ARR(ByVal str_rdx_2 As String, _
                      ByVal str_rdx10 As String, _
                      ByVal str_rdx16 As String)
  Dim i As Long
  
  '探索..空いてる配列に文字列を格納
  For i = LBound(ARR_STR, 1) To UBound(ARR_STR, 1)
    If ARR_STR(i, 1) = "" Then  '値代入
      ARR_STR(i, 0) = str_rdx_2
      ARR_STR(i, 1) = str_rdx10
      ARR_STR(i, 2) = str_rdx16
    End If
    
    '重複を見つけたらカウントアップしてループを抜ける
    If str_rdx10 = ARR_STR(i, 1) Then
      ARR_CNT(i) = ARR_CNT(i) + 1 'COUNT UP
      Exit For
    End If
  Next i
  
End Sub


'---------------------------------------------------------------------------
'配列の並び替え
'  入力文字列の出現頻度を、降順でソートする
'→(配列数-1)!数処理してしまう。14!=800億。処理数が少なるよう改善。2022_0120
'  引数    (IN)    ：なし
'  戻り値  (OUT)   ：なし
'---------------------------------------------------------------------------
Sub SORT_ARR()

  Dim vSwap_str(2) As String '0:2進数, 1:10進数, 2:16進数
  Dim vSwap_cnt As Long
  Dim i, j, k As Long
  
  For i = LBound(ARR_STR, 1) To UBound(ARR_STR, 1)
  
    If (ARR_CNT(i) = 0) Then
      Exit For '探索終了
    End If
    
    If (ARR_CNT(i) < ARR_CNT(i + 1)) Then 'ソート開始
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
      Exit For 'ソート完了
    End If
  Next i

'  'バブルソート  非効率だった
'  For i = UBound(ARR_STR, 1) To LBound(ARR_STR, 1) Step -1 '探索
'    For j = 0 To (i - 1)
'      '大小関係不一致で並び替えを実施
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
'文字列に、指定した文字間隔で区切り文字を挿入
'
'  引数1   (IN)    ：進数指定      0:2進数, 1:10進数, 2:16進数
'  引数2   (IN)    ：区切る間隔
'  引数3   (IN)    ：何で区切るか  "_"/","等
'  戻り値  (OUT)   ：なし
'---------------------------------------------------------------------------
Private Sub SplitStr(ByVal RdxNo As Integer, _
                     ByVal StrLength As Long, _
                     ByVal Char As String)
  
  Dim stt_mid As Long   'MID関数のスタート位置
  Dim o_str As String   '結果格納用
  Dim ModStrLen As Long 'あまりの文字数
  Dim i, j As Long
  
  For i = LBound(ARR_STR_DISP, 1) To UBound(ARR_STR_DISP, 1)
    If (ARR_CNT(i) = 0) Then
      Exit For '終了
    Else
      '初期化
      stt_mid = 1
      o_str = ""
      ModStrLen = Len(ARR_STR_DISP(i, RdxNo)) Mod StrLength
        
      If (ModStrLen <> 0) Then
        stt_mid = stt_mid + ModStrLen
        o_str = Mid(ARR_STR_DISP(i, RdxNo), 1, ModStrLen)
      End If
        
      'xx文字数の間隔で、"_"や","を差し込んでいく
      For j = stt_mid To Len(ARR_STR_DISP(i, RdxNo)) Step StrLength
        If (j = 1) Then
          o_str = o_str & Mid(ARR_STR_DISP(i, RdxNo), j, StrLength)
        Else             '↓"_"や","
          o_str = o_str & Char & Mid(ARR_STR_DISP(i, RdxNo), j, StrLength)
        End If
      Next j
        
      ARR_STR_DISP(i, RdxNo) = o_str
    End If
  Next i
End Sub


'---------------------------------------------------------------------------
'文字列を1文字ずつ分解しながら書き込み
'  引数1   (IN)    ：文字列
'  引数2   (IN)    ：書き込み箇所
'  戻り値  (OUT)   ：なし
'---------------------------------------------------------------------------
Private Sub StringDecomposition(ByVal str As String, ByVal rng As Range)
  Dim i As Long
  
  With rng
    For i = 1 To Len(str) '文字列の長さ分ループ
      .Offset(0, -i).Value = Mid(str, i, 1) '一文字ずつ書き込む
   Next i
  End With

End Sub


'---------------------------------------------------------------------------
'文字色装飾
'  引数1   (IN)    ：変換前の基数
'  引数2   (IN)    ：変更したい文字色
'  戻り値  (OUT)   ：なし
'---------------------------------------------------------------------------
Sub ResultColor(ByVal i_rdx As String, ByVal i_clolor As Long)
  Select Case i_rdx
    Case "2進数"
      Pg_Result_Range().Rows(1).Font.Color = i_clolor ' 2進数
  
    Case "10進数"
      Pg_Result_Range().Rows(2).Font.Color = i_clolor '10進数
    
    Case "16進数"
      Pg_Result_Range().Rows(3).Font.Color = i_clolor '16進数
  
  End Select
End Sub


'---------------------------------------------------------------------------
'検索回数ランキング書き込み
'  引数1   (IN)    ：書き込み開始のセル位置
'  戻り値  (OUT)   ：なし
'---------------------------------------------------------------------------
Private Sub RANKING_WRITE(ByVal stt_rng As Range)
  Dim i, j As Long
  
    '書き込み開始位置(stt_rng)を基準に、検索頻度のランキングを書き込む。
    For i = LBound(ARR_STR_DISP, 1) To UBound(ARR_STR_DISP, 1)   '行方向
      
      If (ARR_CNT(i) = 0) Or (i >= RANK_DISP_NUM_MAX) Then
        Exit For '終了
      Else
        For j = LBound(ARR_STR_DISP, 2) To UBound(ARR_STR_DISP, 2) '列方向
          stt_rng.Offset(i, j).Value = ARR_STR_DISP(i, j)
        Next j
          stt_rng.Offset(i, 3).Value = ARR_CNT(i) & "回"
      End If
    Next i
    
    '文字列長に合わせて、列幅を自動調整
    For i = 2 To 5
      stt_rng.CurrentRegion.Columns(i).AutoFit
    Next i
    
End Sub


'---------------------------------------------------------------------------
'データベースに履歴を書き込む
'  引数1   (IN)    ：変換後の 2進数文字列
'  引数2   (IN)    ：変換後の10進数文字列
'  引数3   (IN)    ：変換後の16進数文字列
'  戻り値  (OUT)   ：なし
'---------------------------------------------------------------------------
Private Sub DATABASE_WRITE(ByVal o_radix__2 As String, _
                           ByVal o_radix_10 As String, _
                           ByVal o_radix_16 As String)
  
  Dim w_rw As Long '書き込む行(数字)
  Dim i As Long
  
  '最終行の1行下の行数を取得
  w_rw = Pg_WSobj_DB.Cells(Rows.count, "B").End(xlUp).Row + 1
  
  Pg_WSobj_DB.Cells(w_rw, "B").Value = o_radix__2 ' 2進数結果
  Pg_WSobj_DB.Cells(w_rw, "C").Value = o_radix_10 '10進数結果
  Pg_WSobj_DB.Cells(w_rw, "D").Value = o_radix_16 '16進数結果
  Pg_WSobj_DB.Cells(w_rw, "E").Value = Date       '実行日
  
End Sub


'---------------------------------------------------------------------------
'メインシート : 変換結果セルをクリア
'  引数1   (IN)    ：なし
'  戻り値  (OUT)   ：なし
'---------------------------------------------------------------------------
Public Sub RESULT_CLR()
  
  '変換前の値入力セルを選択
  If ActiveSheet.Name = Pg_WSName_Main Then
    With Pg_I_DAT
    .Select  '値入力セルにカーソル合わせる
    End With
  End If
  
  '結果表示セルの値をクリア
  With Pg_Result_Range()
    .ClearContents '文字削除
    .Font.ColorIndex = xlAutomatic '文字色を自動(黒)に
  End With
  
End Sub


'---------------------------------------------------------------------------
'データベースシート：変換履歴をクリア
'  引数1   (IN)    ：なし
'  戻り値  (OUT)   ：なし
'---------------------------------------------------------------------------
Public Sub DATABASE_CLR()
  Dim return_msg As VbMsgBoxResult  'VbMsgBoxResult列挙体
  Dim i As Long
  Dim j As Long
  
  return_msg = MsgBox("削除します。宜しいですか？", _
                      vbYesNo, "確認")
  
  If return_msg = vbYes Then
    '配列初期化
    For i = LBound(ARR_STR, 1) To UBound(ARR_STR, 1)   '行方向
      For j = LBound(ARR_STR, 2) To UBound(ARR_STR, 2) '列方向
        ARR_STR(i, j) = ""
      Next j
        ARR_CNT(i) = 0
    Next i
    
    'データベース範囲の値をクリア
    Pg_History_DB_Stt.CurrentRegion.Offset(1, 0).ClearContents
    
    'ランキング表示範囲をクリア
    Pg_Ranking_Main_Stt.CurrentRegion.Offset(1, 1).ClearContents 'メインシート
    'Pg_Ranking_DB_Stt.CurrentRegion.Offset(1, 1).ClearContents   '履歴シート
    
    'カーソル位置の調整
    If ActiveSheet.Name = Pg_WSName_Main Then
      Pg_I_DAT.Select  '値入力セルにカーソル合わせる
    Else
      Range("A1").Select 'A1セル選択
    End If
  End If
  
End Sub
