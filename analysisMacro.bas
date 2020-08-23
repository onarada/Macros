Sub Macro1()
    'ユニーク配列の取得
    Worksheets("データソース").Activate
    Dim sampleArray, koumokuArray
    sampleArray = Distinct(Range(Range("B2"), Cells(Rows.Count, 2).End(xlUp)).Value)
    koumokuArray = Distinct(Range(Range("C2"), Cells(Rows.Count, 3).End(xlUp)).Value)
    For Each koumokuID In koumokuArray
        For Each sampleID In sampleArray
            sheetTitle = koumokuID & "-" & sampleID
            Worksheets("データソース").copy After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = sheetTitle
            With Range("A1")
                .AutoFilter Field:=1, Criteria1:=sampleID
                .AutoFilter Field:=2, Criteria1:=koumokuID
            End With
            'データ貼り付け
            'データソースシートを選択
            Sheets(sheetTitle).Select
            '項目名をフォーマットに貼り付け
            Range(Range("C1"), Cells(Rows.Count, 3).End(xlUp)).copy Sheets(1).Range("A1")
            '濃度をフォーマットに貼り付け
            Range(Range("D1"), Cells(Rows.Count, 4).End(xlUp)).copy Sheets(1).Range("B1")
            'BASをフォーマットに貼り付け
            Range(Range("G1:I1"), Cells(Rows.Count, 7).End(xlUp)).copy Sheets(1).Range("C1")
            'BAMををフォーマットに貼り付け
            Range(Range("J1:L1"), Cells(Rows.Count, 10).End(xlUp)).copy Sheets(1).Range("C12")
            'RASをフォーマットに貼り付け
            Range(Range("N1:AY1"), Cells(Rows.Count, 14).End(xlUp)).copy Sheets(1).Range("F1")
            'RAMをフォーマットに貼り付け
            Range(Range("BA1:CL1"), Cells(Rows.Count, 53).End(xlUp)).copy Sheets(1).Range("AR1")
            '解析データシートを複製
            Sheets(1).copy After:=Sheets(Sheets.Count)
            '解析データシート名に項目名とサンプル名を挿入
            sheetname = koumokuID & "×" & sampleID
            ActiveSheet.Name = sheetname
            '解析データシートの図形を削除
            ActiveSheet.Shapes.SelectAll
            Selection.Delete
            'データ貼り付け元シートを削除
            Sheets(sheetTitle).Select
            Application.DisplayAlerts = False
            ActiveSheet.Delete
            'フォーマットのデータをクリア
            Sheets(1).Select
            Range("A2:CC11").Clear
            Range("C13:E22").Clear
        Next sampleID
    Next koumokuID
End Sub

Function Distinct(args As Variant) As Variant

    Dim dictionary   As Object
    Set dictionary = CreateObject("scripting.dictionary")

   'microsoft scripting runtimeが参照設定されている場合は以下のほうが良い(補完が効く上に多少速い)
   'Dim dictionary   As Dictionary
   'set dictionary = new Dictionary

    Dim arg As Variant
    For Each arg In args
        If Not dictionary.Exists(arg) Then
            dictionary.Add arg, 0
        End If
    Next arg

    Distinct = dictionary.Keys

End Function
