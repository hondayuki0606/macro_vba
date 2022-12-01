<pre>
Sub MacroUpdateAN()
        
    Dim rc As VbMsgBoxResult
    rc = MsgBox("マクロ実行前に必ずファイルのバックアップを行ってください。マクロを実行しますか。", vbYesNo + vbDefaultButton2)
    
    If rc = vbNo Then
        MsgBox "マクロの処理をキャンセルしました"
    Else
        'AN列のインデックス番号
        Const C_AN_COLUMN As Long = 40
        'AO列のインデックス番号
        Const C_AO_COLUMN As Long = 41
        
        'AO列の最終行番号を取得
        Dim MaxAOColumn As Long
        MaxAOColumn = Range("AO1").SpecialCells(xlLastCell).Row
    
        'MsgBox "AO列の最終行は" & MaxAOColumn & "行目"
        
        ' ★連番の初期値を変えたい場合はSeqNo、C_SEQ_NOの2つを変更する★
        Dim SeqNo As Long
        SeqNo = 1
        Const C_SEQ_NO As Long = 1
        
        
        Dim TargetRowNo As Long
        Dim NextRowNo As Long
        
        ' AN列を初期化
        For Row = 1 To MaxAOColumn
            Cells(Row, C_AN_COLUMN) = ""
        Next Row
        
        ' AN列の書き込み処理
        For Row = 1 To MaxAOColumn
        
            ' シーケンスが初期値(C_SEQ_NO:1)の場合 ※初回書き込みのみ実行される処理
            If SeqNo = C_SEQ_NO Then
                ' AOに0か空文字の場合は、処理をスキップ
                If Cells(Row, C_AO_COLUMN) = "" Or Cells(Row, C_AO_COLUMN) = "0" Then
                    GoTo CONTINUE:
                End If
                ' 値が有効値の場合シーケンスを書き込む
                Cells(Row, C_AN_COLUMN) = SeqNo
                SeqNo = SeqNo + 1
                GoTo CONTINUE:
            End If
            
            
            ' AOのターゲットの値が空文字の場合、処理をスキップする
            If Cells(Row, C_AO_COLUMN) = "" Then
                GoTo CONTINUE:
            End If
            
            ' AOのターゲットの値が0の場合
            If Cells(Row, C_AO_COLUMN) = "0" Then
                ' AOの次の行が空文字の場合、処理をスキップする
                If Cells(Row + 1, C_AO_COLUMN) = "" Then
                    GoTo CONTINUE:
                End If
                ' AOのターゲットの値と次の行の値を比較し、次の行の値の方が小さい場合、次の行のAN列にシーケンスを書き込む
                If TargetRowNo > Cells(Row + 1, C_AO_COLUMN) Then
                    Cells(Row + 1, C_AN_COLUMN) = SeqNo
                    SeqNo = SeqNo + 1
                End If
                GoTo CONTINUE:
            End If
            TargetRowNo = Cells(Row, C_AO_COLUMN)
            
            ' AOの次の行の値が空文字 または 0の場合、処理をスキップする
            If Cells(Row + 1, C_AO_COLUMN) = "" Or Cells(Row + 1, C_AO_COLUMN) = "0" Then
                GoTo CONTINUE:
            End If
            NextRowNo = Cells(Row + 1, C_AO_COLUMN)
            
            ' AOのターゲットの値と次の行の値を比較し、次の行の値の方が小さい場合、次の行のAN列にシーケンスを書き込む
            If TargetRowNo > NextRowNo Then
                Cells(Row + 1, C_AN_COLUMN) = SeqNo
                SeqNo = SeqNo + 1
            End If
CONTINUE:
        Next Row
        MsgBox "AN列の更新が完了しました。"
    End If
End Sub

</pre>
