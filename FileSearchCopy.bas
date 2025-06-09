Option Explicit

Sub SearchAndCopyFiles()
    Dim keyword As String
    keyword = InputBox("請輸入要搜尋的關鍵字：（不分大小寫）")
    If keyword = "" Then Exit Sub

    Dim defaultExt As String
    defaultExt = "*.ppt;*.doc;*.docx;*.xls;*.xlsm;*.pdf;*.txt;*.pptx"

    Dim extInput As String
    extInput = InputBox("請輸入要搜尋的檔案格式，以分號;分隔:\n例如: *.ppt;*.docx;*.pdf", "檔案格式", defaultExt)
    If extInput = "" Then Exit Sub

    Dim srcFolder As String
    srcFolder = GetFolder("選取來源資料夾")
    If srcFolder = "" Then Exit Sub

    Dim destFolder As String
    destFolder = GetFolder("選取目的資料夾")
    If destFolder = "" Then Exit Sub

    Dim exts() As String
    exts = Split(extInput, ";")

    Dim sht As Worksheet
    Set sht = ActiveSheet
    sht.Range("A3:C" & sht.Rows.Count).Clear

    Dim rowNum As Long
    rowNum = 3

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    SearchAndCopy fso.GetFolder(srcFolder), keyword, exts, destFolder, sht, rowNum
    MsgBox "完成！共複製 " & (rowNum - 3) & " 個檔案。"
End Sub

Function GetFolder(prompt As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = prompt
    If fd.Show = -1 Then
        GetFolder = fd.SelectedItems(1)
    Else
        GetFolder = ""
    End If
End Function

Sub SearchAndCopy(ByVal folder As Object, keyword As String, exts() As String, destFolder As String, sht As Worksheet, ByRef rowNum As Long)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    For Each file In folder.Files
        If MatchExtension(file.Name, exts) Then
            If InStr(1, file.Name, keyword, vbTextCompare) > 0 Then
                fso.CopyFile file.Path, destFolder & "\" & file.Name, True
                sht.Cells(rowNum, 1).Value = rowNum - 2
                sht.Cells(rowNum, 2).Value = file.Name
                sht.Hyperlinks.Add Anchor:=sht.Cells(rowNum, 3), Address:=file.Path, TextToDisplay:=file.Path
                rowNum = rowNum + 1
            End If
        End If
    Next file
    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        SearchAndCopy subFolder, keyword, exts, destFolder, sht, rowNum
    Next subFolder
End Sub

Function MatchExtension(fileName As String, exts() As String) As Boolean
    Dim i As Integer, e As String, ext As String
    ext = LCase$(Mid$(fileName, InStrRev(fileName, ".") + 1))
    For i = LBound(exts) To UBound(exts)
        e = Trim(exts(i))
        If e <> "" Then
            e = LCase$(Replace(e, "*.", ""))
            If ext = e Then
                MatchExtension = True
                Exit Function
            End If
        End If
    Next i
    MatchExtension = False
End Function

Sub AddSearchButton()
    Dim btn As Button
    On Error Resume Next
    ActiveSheet.Buttons("btnSearchCopy").Delete
    On Error GoTo 0
    Set btn = ActiveSheet.Buttons.Add(ActiveSheet.Range("A2").Left, _
                                      ActiveSheet.Range("A2").Top, _
                                      ActiveSheet.Range("A2").Width, _
                                      ActiveSheet.Range("A2").Height)
    With btn
        .Caption = "搜尋並複製檔案"
        .OnAction = "SearchAndCopyFiles"
        .Name = "btnSearchCopy"
    End With
End Sub
