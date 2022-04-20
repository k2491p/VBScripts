Call Main

'----------------------------------------------------
'
'   【Excelフォーマット】
'   本ファイルがあるディレクトリと同じディレクトリにある
'   すべてのExcel(*.xlsx, *.xlsm, *.xls)に対して、
'       ・A1をセルを選択
'       ・拡大率を100%
'   を全シートに行います。
'
'----------------------------------------------------
Sub Main()
    WScript.Echo "Excelフォーマット処理開始"
    'Excelインスタンス生成
    Dim objXlsx : Set objXlsx = CreateObject("Excel.Application")
    If IsNull(objXlsx) Then 
        Exit Sub
    End If

    'Excel非表示
    objXlsx.Visible = False
    
    '上書き保存のアラート非表示
    objXlsx.DisplayAlerts = False

    'FileSystemObjectインスタンス生成
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

    'カレントディレクトリをセット
    Dim currentDirectory : Set currentDirectory = fso.GetFolder(".\")

    'カウント数 (カレントディレクトリにvbsとbatがある想定でマイナス2)
    Dim fileCount : fileCount = currentDirectory.Files.Count - 2
    Dim count : count = 1
    
    'ディレクトリ内を処理
    For Each file In currentDirectory.Files
        If IsExcel(fso.GetExtensionName(file)) Then
            Dim workbook : Set workbook = objXlsx.Workbooks.Open(file)
            SetAllA1(workbook)
            workbook.Saveas(file)
            workbook.Close
            Set workbook = Nothing
            WScript.Echo fileCount & "件中" & count & "件目処理終了"
            count = count + 1 
        End If
    Next
    objXlsx.Quit()
End Sub

'Excelかどうか判定
Function IsExcel (extention)
    IsExcel = False
    If extention = "xlsx" Or extention = "xls" Or extention = "xlsm" then
        IsExcel = True
    End If
End Function

'引数のブック内をすべてA1にセット
Sub SetAllA1 (workbook)
    Dim worksheet
    For Each worksheet In workbook.Worksheets
        workbook.Worksheets(worksheet.Name).Activate
        worksheet.Range("A1").Activate
        workbook.Windows(1).ScrollRow = 1
        workbook.Windows(1).ScrollColumn = 1
        workbook.Windows(1).Zoom = 100
    Next
    workbook.WorkSheets(1).Activate
End Sub