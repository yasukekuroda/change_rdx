Attribute VB_Name = "基数変換_A"
Option Explicit
'---------------------------------------------------------------------------
'基数変換管理 ..A方法
'  参考サイト
'  Excel作業をVBAで効率化
'  https://vbabeginner.net/convert-hextodec/
'
'
'  引数1   (IN)    ：変換前の基数
'  引数2   (IN)    ：入力値
'  引数3   (IN)    ：変換後の基数 "2進数" or "10進数" or "16進数"
'  戻り値  (OUT)   ：変換した文字列
'---------------------------------------------------------------------------
Public Function RDX_CHANGE_A(ByVal i_rdx As String, _
                             ByVal i_dat As String, _
                             ByVal o_rdx As String)
  Select Case i_rdx
    Case "2進数"
      If o_rdx = "2進数" Then
        RDX_CHANGE_A = i_dat           '2 →  2
      ElseIf o_rdx = "10進数" Then
        RDX_CHANGE_A = BinToDec(i_dat) '2 → 10
      Else
        RDX_CHANGE_A = BinToHex(i_dat) '2 → 16
      End If
      
    Case "10進数"
      If o_rdx = "2進数" Then
        RDX_CHANGE_A = DecToBin(i_dat) '10 →  2
      ElseIf o_rdx = "10進数" Then
        RDX_CHANGE_A = i_dat           '10 → 10
      Else
        RDX_CHANGE_A = Hex(i_dat)      '10 → 16
      End If
    
    Case "16進数"
      If o_rdx = "2進数" Then
        RDX_CHANGE_A = DecToBin(CStr(HexToDec(i_dat))) ' 16 → 10 → 2
      ElseIf o_rdx = "10進数" Then
        RDX_CHANGE_A = HexToDec(i_dat)                  '16 → 10
      Else
        RDX_CHANGE_A = UCase(i_dat)                     '16 → 16
      End If
      
  End Select
End Function


'---------------------------------------------------------------------------
'基数変換(2→10)
'  引数1   (IN)    ：2進数文字列
'  戻り値  (OUT)   ：10進数文字列
'---------------------------------------------------------------------------
Private Function BinToDec(ByVal I_BIN As String)
  Dim i        As Long
  Dim i_Len    As Long    '  2進数文字列長
  Dim sParts   As String   '  2進数文字列を切り出した一部
  Dim O_DEC    As Long    ' 10進数値
    
  O_DEC = 0 '初期化

  i_Len = Len(I_BIN) '入力の文字列長を取得
    
  For i = 1 To i_Len
    '2進数文字列を1文字切り出し
    sParts = Mid(I_BIN, i, 1)
        
    '2進数値×2のn乗の値を加算する
    O_DEC = O_DEC + 2 ^ (i_Len - i) * CLng(sParts)
  Next i
    
  '文字列として出力
  BinToDec = CStr(O_DEC)
End Function


'---------------------------------------------------------------------------
'基数変換(2→16)
'  引数1   (IN)    ：2進数文字列
'  戻り値  (OUT)   ：16進数文字列
'---------------------------------------------------------------------------
Private Function BinToHex(ByVal I_BIN As String)
  Dim i           As Long
  Dim iParts      As Long    '  2進数文字列ループカウンタ
  Dim sParts      As String  '  2進数文字列を切り出した4文字
  Dim iRemainder  As Long    ' 余り
  Dim iDec        As Long    ' 10進数値
    
  '2進数文字列の文字数を4で割った余りを取得
  iRemainder = Len(I_BIN) Mod 4
    
  '余りがある場合、不足している"0"を付与
  If (iRemainder > 0) Then
    I_BIN = Left("0000", 4 - iRemainder) & I_BIN
  End If
    
  '2進数文字列を左から4文字ずつループ
  For iParts = 1 To Len(I_BIN) Step 4
    '2進数文字列を4文字切り出し
    sParts = Mid(I_BIN, iParts, 4)
        
    '10進数値を初期化
    iDec = 0
        
    '切り出した文字列を左から順に1文字ずつ10進数値に変換して4文字分を合計する
    For i = 0 To 3
      '2進数値×2のn乗の値を加算する
      iDec = iDec + 2 ^ (3 - i) * CInt(Mid(sParts, i + 1, 1))
    Next i
        
    '0000以外の入力に対して実行
    If (BinToHex <> "" Or iDec <> 0) Then
      '16進数文字列を連結
      BinToHex = BinToHex & CStr(Hex(iDec))
    End If
  Next iParts
    
  '1'b0の入力に対しては0を出力
  If (BinToHex = "") Then
      BinToHex = "0"
  End If

End Function


'---------------------------------------------------------------------------
'基数変換(10→2)
'  引数1   (IN)    ：10進数文字列
'  戻り値  (OUT)   ： 2進数文字列
'---------------------------------------------------------------------------
Public Function DecToBin(ByVal a_sDec As String)
  Dim i           As Long
  Dim iRemainder  As Long    '余り
  Dim dDiv        As Double  '商
    
  '引数10進数文字列を10進数値として取得
  dDiv = Val(a_sDec)
    
  '処理完了までループ
  Do
    '10進数値を2で割った余りを取得
    iRemainder = dDiv Mod 2
    
    '10進数値を2で割った商を取得（次ループの10進数値になる）
    dDiv = Int(dDiv / 2)
        
    '2進数文字列の左に余りを連結
    DecToBin = CStr(iRemainder) & DecToBin
        
    '10進数値が2未満（もう2で割れないのでここでループ終了）
    If (dDiv < 2) Then
      If (dDiv = 1) Then
        DecToBin = CStr(dDiv) & DecToBin '最上位桁の値として"1"を連結
      End If
      Exit Do 'ループを抜ける
    End If
  Loop

End Function


'---------------------------------------------------------------------------
'基数変換(16→10)
'  引数1   (IN)    ：16進数文字列
'  戻り値  (OUT)   ：10進数文字列
'---------------------------------------------------------------------------
Private Function HexToDec(ByVal a_sHex As String)
  Dim dDec As Double   '10進数値
    
  '10進数値に変換
  dDec = CLng("&H" & a_sHex)
    
  HexToDec = CStr(dDec)
End Function
