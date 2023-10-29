Attribute VB_Name = "Module1"
'#############################################################################
' CSVファイルをTXTファイルに変換
'　convert_csv2txt
'#############################################################################
Sub ConvertCSVtoTXT()
    Dim OutputCsv As String, OutputTxt As String
    Dim InputFileNumber As Integer, OutputFileNumber As Integer
    Dim LineText As String
    ' CSVファイルのパスを設定
    OutputCsv = Environ("USERPROFILE") & "\Desktop\sample.csv"
    ' TXTファイルのパスを設定
    OutputTxt = Environ("USERPROFILE") & "\Desktop\sample.txt"
    ' CSVファイルをテキストファイルに変換
    InputFileNumber = FreeFile
    Open OutputCsv For Input As #InputFileNumber
    OutputFileNumber = FreeFile
    Open OutputTxt For Output As #OutputFileNumber
    Do While Not EOF(InputFileNumber)
        Line Input #InputFileNumber, LineText
        ' ダブルクォーテーション3つ連続をダブルクォーテーション1つに置換
        LineText = Replace(LineText, """""", "")
        ' 行末の空白を削除
        LineText = Trim(LineText)
        ' テキストファイルに書き込み
        Print #OutputFileNumber, LineText
    Loop
    Close #InputFileNumber
    Close #OutputFileNumber
    MsgBox "CSVファイルをテキストファイルに変換しました。", vbInformation
End Sub
