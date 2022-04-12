Call Main

'----------------------------------------------------
'
'   �yExcel�t�H�[�}�b�g�z
'   �{�t�@�C��������f�B���N�g���Ɠ����f�B���N�g���ɂ���
'   ���ׂĂ�Excel(*.xlsx, *.xlsm, *.xls)�ɑ΂��āA
'       �EA1���Z����I��
'       �E�g�嗦��100%
'   ��S�V�[�g�ɍs���܂��B
'
'----------------------------------------------------
Sub Main()
    WScript.Echo "Excel�t�H�[�}�b�g�����J�n"
    'Excel�C���X�^���X����
    Dim objXlsx : Set objXlsx = CreateObject("Excel.Application")
    If IsNull(objXlsx) Then 
        Exit Sub
    End If

    'Excel��\��
    objXlsx.Visible = False
    
    '�㏑���ۑ��̃A���[�g��\��
    objXlsx.DisplayAlerts = False

    'FileSystemObject�C���X�^���X����
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

    '�J�����g�f�B���N�g�����Z�b�g
    Dim currentDirectory : Set currentDirectory = fso.GetFolder(".\")

    '�J�E���g�� (�J�����g�f�B���N�g����vbs��bat������z��Ń}�C�i�X2)
    Dim fileCount : fileCount = currentDirectory.Files.Count - 2
    Dim count : count = 1
    
    '�f�B���N�g����������
    For Each file In currentDirectory.Files
        If IsExcel(fso.GetExtensionName(file)) Then
            Dim workbook : Set workbook = objXlsx.Workbooks.Open(file)
            SetAllA1(workbook)
            workbook.Saveas(file)
            workbook.Close
            Set workbook = Nothing
            WScript.Echo fileCount & "����" & count & "���ڏ����I��"
            count = count + 1 
        End If
    Next
    objXlsx.Quit()
End Sub

'Excel���ǂ�������
Function IsExcel (extention)
    IsExcel = False
    If extention = "xlsx" Or extention = "xls" Or extention = "xlsm" then
        IsExcel = True
    End If
End Function

'�����̃u�b�N�������ׂ�A1�ɃZ�b�g
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