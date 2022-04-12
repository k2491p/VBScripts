Call Main

'----------------------------------------------------
'
'   ï¿½yExcelï¿½tï¿½Hï¿½[ï¿½}ï¿½bï¿½gï¿½z
'   ï¿½{ï¿½tï¿½@ï¿½Cï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½fï¿½Bï¿½ï¿½ï¿½Nï¿½gï¿½ï¿½ï¿½Æ“ï¿½ï¿½ï¿½ï¿½fï¿½Bï¿½ï¿½ï¿½Nï¿½gï¿½ï¿½ï¿½É‚ï¿½ï¿½ï¿½
'   ï¿½ï¿½ï¿½×‚Ä‚ï¿½Excel(*.xlsx, *.xlsm, *.xls)ï¿½É‘Î‚ï¿½ï¿½ÄA
'       ï¿½EA1ï¿½ï¿½Zï¿½ï¿½ï¿½ï¿½Iï¿½ï¿½
'       ï¿½Eï¿½gï¿½å—¦ï¿½ï¿½100%
'   ï¿½ï¿½Sï¿½Vï¿½[ï¿½gï¿½Ésï¿½ï¿½ï¿½Ü‚ï¿½ï¿½B
'
'----------------------------------------------------
Sub Main()
    WScript.Echo "ExcelƒtƒH[ƒ}ƒbƒgˆ—ŠJn"
    'ExcelƒCƒ“ƒXƒ^ƒ“ƒX¶¬
    Dim objXlsx : Set objXlsx = CreateObject("Excel.Application")
    If IsNull(objXlsx) Then 
        Exit Sub
    End If

    'Excel”ñ•\¦
    objXlsx.Visible = False
    
    'ã‘‚«•Û‘¶‚ÌƒAƒ‰[ƒg”ñ•\¦
    objXlsx.DisplayAlerts = False

    'FileSystemObjectƒCƒ“ƒXƒ^ƒ“ƒX¶¬
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

    'ƒJƒŒƒ“ƒgƒfƒBƒŒƒNƒgƒŠ‚ğƒZƒbƒg
    Dim currentDirectory : Set currentDirectory = fso.GetFolder(".\")

    'ƒJƒEƒ“ƒg”
    Dim fileCount : fileCount = currentDirectory.Files.Count
    Dim count : count = 1
    
    'ƒfƒBƒŒƒNƒgƒŠ“à‚ğˆ—
    For Each file In currentDirectory.Files
        If IsExcel(fso.GetExtensionName(file)) Then
            Dim workbook : Set workbook = objXlsx.Workbooks.Open(file)
            SetAllA1(workbook)
            workbook.Saveas(file)
            workbook.Close
            Set workbook = Nothing
        End If
        WScript.Echo fileCount & "Œ’†" & count & "Œ–Úˆ—I—¹"
        count = count + 1 
    Next
    objXlsx.Quit()
End Sub

'Excel‚©‚Ç‚¤‚©”»’è
Function IsExcel (extention)
    IsExcel = False
    If extention = "xlsx" Or extention = "xls" Or extention = "xlsm" then
        IsExcel = True
    End If
End Function

'ˆø”‚ÌƒuƒbƒN“à‚ğ‚·‚×‚ÄA1‚ÉƒZƒbƒg
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