Attribute VB_Name = "基数変換_B"
Option Explicit

Private i As Long
Private varBinary As Variant
Private colHValue As New Collection '連想配列、Collectionオブジェクトの作成
Private lngNu() As Long


'---------------------------------------------------------------------------
'基数変換管理 ..B方法
'引用
'  文字列←→16進数←→2進数の相互変換
'  https://excel.syogyoumujou.com/memorandum/hex_binary.html
'  二進数から十進数を得るユーザ定義関数
'  https://www.moug.net/tech/exvba/0100013.html
'
'  引数1   (IN)    ：変換前の基数
'  引数2   (IN)    ：入力値
'  引数3   (IN)    ：変換後の基数 "2進数" or "10進数" or "16進数"
'  戻り値  (OUT)   ：変換した文字列
'---------------------------------------------------------------------------
'文字列16進数←→2進数の相互変換
Public Function RDX_CHANGE_B(ByVal i_rdx As String, _
                             ByVal i_dat As String, _
                             ByVal o_rdx As String)
                             
  Dim strData As String
  
  '2←→16 進数の変換をリストで対応づける。
  '変換結果で不要な0をトリミングし、出力。
  varBinary = Array("0000", "0001", "0010", "0011", _
                    "0100", "0101", "0110", "0111", _
                    "1000", "1001", "1010", "1011", _
                    "1100", "1101", "1110", "1111")
                    
  Set colHValue = New Collection '初期化
  
  '連想配列に「キー」としてvarBinaryの2進数、「アイテム」として対応する16進数「0～F」を格納
  For i = 0 To 15
    colHValue.Add CStr(Hex$(i)), varBinary(i)
  Next
    
  Select Case i_rdx
    Case "2進数"
      If o_rdx = "2進数" Then
        RDX_CHANGE_B = i_dat                  '2 →  2
      ElseIf o_rdx = "10進数" Then
        RDX_CHANGE_B = SampleBinToDeci(i_dat) '2 → 10
      Else
        RDX_CHANGE_B = BtoH(i_dat)            '2 → 16
      End If
      
    Case "10進数"
      If o_rdx = "2進数" Then
        RDX_CHANGE_B = ExDeciToBin(CLng(i_dat)) '10 →  2
      ElseIf o_rdx = "10進数" Then
        RDX_CHANGE_B = i_dat                    '10 → 10
      Else
        RDX_CHANGE_B = Hex(i_dat)               '10 → 16
      End If
    
    Case "16進数"
      If o_rdx = "2進数" Then
        RDX_CHANGE_B = Trimming(HtoB(i_dat))    '16 →  2
      ElseIf o_rdx = "10進数" Then
        RDX_CHANGE_B = HexToDec(i_dat)          '16 → 10
      Else
        RDX_CHANGE_B = UCase(i_dat)             '16 → 16
      End If
      
  End Select
  
  Erase lngNu '配列の解放...なくても解放される

End Function


'---------------------------------------------------------------------------
'基数変換(16→2)
'  引数1   (IN)    ：16進数文字列
'  戻り値  (OUT)   ： 2進数文字列
'---------------------------------------------------------------------------
Private Function HtoB(ByVal strH As String) As String '16進数→2進数
    ReDim strHtoB(1 To Len(strH)) As String
    For i = 1 To Len(strH)
        strHtoB(i) = varBinary(Val("&h" & Mid$(strH, i, 1)))
    Next
    HtoB = Join$(strHtoB, vbNullString)
End Function


'---------------------------------------------------------------------------
'基数変換(2→16)
'  引数1   (IN)    ： 2進数文字列
'  戻り値  (OUT)   ：16進数文字列
'---------------------------------------------------------------------------
Private Function BtoH(ByVal strB As String) As String '2進数→16進数
  '文字数を4の倍数に調整
  If ((Len(strB) Mod 4) > 0) Then
    strB = Left("0000", 4 - (Len(strB) Mod 4)) & strB
  End If
  
  ReDim strBtoH(1 To Len(strB) / 4) As String
  For i = 1 To Len(strB) / 4 '2進数(4bit分)を16進数に変換
    strBtoH(i) = colHValue.Item(Mid$(strB, (i - 1) * 4 + 1, 4))
  Next
  BtoH = Join$(strBtoH, vbNullString)
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
    i_dat = "0"
  End If
  
  Trimming = i_dat
End Function

'最大の２のべき乗の値を探す
Private Function Ex2noBeki(deci As Long) As Integer
    Dim i As Integer
    
    i = 0
    Do
        'deciより大きい
        If deci < 2 ^ i Then
            'その一つ前のべき乗
            Ex2noBeki = i - 1
            Exit Function
        End If
        i = i + 1
    Loop
End Function

'---------------------------------------------------------------------------
'基数変換(10→2)
'  引数1   (IN)    ：10進数数値
'  戻り値  (OUT)   ： 2進数文字列
'---------------------------------------------------------------------------
Private Function ExDeciToBin(deci As Long) As String
    Dim ln As Long
    Dim stemp As String
    Dim i As Integer
    Dim count As Integer
    
    stemp = "1"
    'deciより小さい、最大の２のべき乗の値を探す
    count = Ex2noBeki(deci)
    ln = deci - 2 ^ count
    '筆算と同じように繰り返す
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
'基数変換(2→10)
'  引数1   (IN)    ： 2進数数値
'  戻り値  (OUT)   ：10進数文字列
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
'この関数だけ 基数変換_Aシートと同じであるが、 Clng()関数を利用している。
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
