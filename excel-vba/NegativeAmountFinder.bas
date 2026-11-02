Attribute VB_Name = "NegativeAmountFinder"
Option Explicit

' 사진 속 "음수찾기" 매크로를 재사용 가능하도록 정리한 버전입니다.
' 지정 폴더와 하위 폴더의 특정 Excel 파일에서 모든 시트의 A열 음수를 찾습니다.
Private resultSheet As Worksheet
Private resultRow As Long
Private skippedFiles As Long

Public Sub 음수찾기()
    Dim fd As FileDialog
    Dim fso As Object
    Dim originalScreenUpdating As Boolean
    Dim originalDisplayAlerts As Boolean

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    If fd.Show <> -1 Then Exit Sub

    originalScreenUpdating = Application.ScreenUpdating
    originalDisplayAlerts = Application.DisplayAlerts
    On Error GoTo HandleError

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set resultSheet = ThisWorkbook.Worksheets.Add
    resultSheet.Range("A1:E1").Value = Array("폴더", "파일", "시트", "셀", "금액")
    resultSheet.Columns("A:D").NumberFormat = "@"
    resultRow = 2
    skippedFiles = 0

    Set fso = CreateObject("Scripting.FileSystemObject")
    폴더검사 fso.GetFolder(fd.SelectedItems(1))

    resultSheet.Columns("A:D").AutoFit
    resultSheet.Columns("E").NumberFormat = "#,##0.############"
    MsgBox "음수 " & (resultRow - 2) & "개 찾음" & vbCrLf & _
           "열지 못했거나 처리 중 오류가 난 파일: " & skippedFiles & "개", vbInformation

CleanUp:
    Application.DisplayAlerts = originalDisplayAlerts
    Application.ScreenUpdating = originalScreenUpdating
    Exit Sub

HandleError:
    MsgBox "폴더 검사 중 오류: " & Err.Description, vbExclamation
    Resume CleanUp
End Sub

Private Sub 폴더검사(ByVal fld As Object)
    Dim fl As Object
    Dim sf As Object

    For Each fl In fld.Files
        ' 사진 속 대상 파일명 조건. 다른 프로젝트에서는 이 패턴을 변경하세요.
        If fl.Name Like "*입출금내역제공내역*.xls*" Then
            If Left$(fl.Name, 2) <> "~$" Then
                파일검사 fl
            End If
        End If
    Next fl

    For Each sf In fld.SubFolders
        폴더검사 sf
    Next sf
End Sub

Private Sub 파일검사(ByVal fl As Object)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim c As Range
    Dim v As Variant

    On Error GoTo FileError

    Set wb = Workbooks.Open(Filename:=fl.Path, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)

    For Each ws In wb.Worksheets
        Set rng = Intersect(ws.UsedRange, ws.Columns("A"))
        If Not rng Is Nothing Then
            For Each c In rng.Cells
                v = c.Value2
                If Not IsError(v) Then
                    If Len(Trim$(CStr(v))) > 0 Then
                        If IsNumeric(v) Then
                            If CDbl(v) < 0 Then
                                resultSheet.Cells(resultRow, 1).Value = fl.ParentFolder.Path
                                resultSheet.Cells(resultRow, 2).Value = fl.Name
                                resultSheet.Cells(resultRow, 3).Value = ws.Name
                                resultSheet.Cells(resultRow, 4).Value = c.Address(False, False)
                                resultSheet.Cells(resultRow, 5).Value = CDbl(v)
                                resultRow = resultRow + 1
                            End If
                        End If
                    End If
                End If
            Next c
        End If
    Next ws

CloseFile:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Sub

FileError:
    skippedFiles = skippedFiles + 1
    Resume CloseFile
End Sub
