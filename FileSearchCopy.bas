Option Explicit

Sub SearchAndCopyFiles()
    Dim keyword As String
    keyword = InputBox("請輸入要搜尋的關鍵字：（不分大小寫）")
    If keyword = "" Then Exit Sub

    Dim exts() As String
    exts = SelectExtensions()
    If Not IsArray(exts) Then Exit Sub

    Dim srcFolder As String
    srcFolder = GetFolder("選取來源資料夾")
    If srcFolder = "" Then Exit Sub

    Dim destFolder As String
    destFolder = GetFolder("選取目的資料夾")
    If destFolder = "" Then Exit Sub


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

Function SelectExtensions() As Variant
    Dim types As Variant
    types = Array("*.ppt", "*.doc", "*.docx", "*.xls", "*.xlsm", "*.pdf", "*.txt", "*.pptx")

    Dim frm As Object
    Set frm = VBA.UserForms.Add
    frm.Caption = "選擇檔案格式"

    Dim chk() As Object
    ReDim chk(UBound(types))
    Dim i As Integer, topPos As Double
    topPos = 10
    For i = LBound(types) To UBound(types)
        Set chk(i) = frm.Controls.Add("Forms.CheckBox.1", "chk" & i)
        chk(i).Caption = types(i)
        chk(i).Left = 10
        chk(i).Top = topPos
        chk(i).Width = 100
        chk(i).Value = True
        topPos = topPos + 18
    Next i

    Dim btnOK As Object, btnCancel As Object
    Set btnOK = frm.Controls.Add("Forms.CommandButton.1", "btnOK")
    btnOK.Caption = "確定"
    btnOK.Left = 10
    btnOK.Top = topPos + 10
    btnOK.Width = 50

    Set btnCancel = frm.Controls.Add("Forms.CommandButton.1", "btnCancel")
    btnCancel.Caption = "取消"
    btnCancel.Left = 70
    btnCancel.Top = topPos + 10
    btnCancel.Width = 50

    With frm.CodeModule
        Dim baseLine As Long
        baseLine = .CountOfLines + 1
        .InsertLines baseLine, "Private mCancel As Boolean" & vbCrLf & _
            "Private Sub btnCancel_Click()" & vbCrLf & _
            "    mCancel = True" & vbCrLf & _
            "    Me.Hide" & vbCrLf & _
            "End Sub" & vbCrLf & _
            "Private Sub btnOK_Click()" & vbCrLf & _
            "    Me.Hide" & vbCrLf & _
            "End Sub" & vbCrLf & _
            "Public Property Get Canceled() As Boolean" & vbCrLf & _
            "    Canceled = mCancel" & vbCrLf & _
            "End Property"
    End With

    frm.Show vbModal

    If VBA.CallByName(frm, "Canceled", VbMethod) Then
        SelectExtensions = Empty
    Else
        Dim coll As New Collection
        For i = LBound(types) To UBound(types)
            If chk(i).Value = True Then
                coll.Add types(i)
            End If
        Next i
        If coll.Count = 0 Then
            SelectExtensions = Empty
        Else
            Dim result() As String
            ReDim result(0 To coll.Count - 1)
            For i = 1 To coll.Count
                result(i - 1) = coll(i)
            Next i
            SelectExtensions = result
        End If
    End If
    Unload frm
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
