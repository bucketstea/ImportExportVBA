Attribute VB_Name = "ImportSrc"
Option Explicit
Public Sub importAll()
    Dim targetFile   As String:   targetFile = ActiveSheet.Cells(6, 2).Value
    Dim targetFolder As String: targetFolder = ActiveSheet.Cells(7, 2).Value
    Call ImportVbaSourcesFromFolder(targetFile, targetFolder)
End Sub

' === 公開エントリポイント（PowerShellから叩く）===
' 引数 folderPath: 展開したソース格納フォルダ（末尾\ ありなしOK）
Private Sub ImportVbaSourcesFromFolder(ByVal filePath As String, _
                                       ByVal folderPath As String)
    Dim wb As Workbook: Set wb = Workbooks.Open(Filename:=filePath, ReadOnly:=False)
    Dim vbProj As VBIDE.VBProject: Set vbProj = wb.VBProject
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        Err.Raise vbObjectError + 1, , "Folder not found: " & folderPath
    End If
    Dim fld As Object: Set fld = fso.GetFolder(folderPath)

    ' 1) 既存モジュールを必要に応じて削除（安全側に「標準/クラスのみ」削除）
'    RemoveAllStdAndClassModules vbProj

    ' 2) ファイルを全インポート
    Dim fil As Object
    For Each fil In fld.Files
        Dim ext As String: ext = LCase$(fso.GetExtensionName(fil.Path))
        Select Case ext
            Case "bas", "cls", "frm"
                vbProj.VBComponents.Import fil.Path
            Case Else
                ' ignore
        End Select
    Next fil
    
    ' 3) フォーム(.frm)がある場合、同名の.frxが必要（同フォルダに置かれていればOK）
    ' 4) 保存
'    ThisWorkbook.Save
    
    MsgBox wb.Name & "に、" & folderPath & "のソースファイルを取り込みました。"
End Sub

Private Sub RemoveAllStdAndClassModules(ByVal vbProj As VBIDE.VBProject)
    Dim comp As VBIDE.VBComponent
    Dim toRemove As Collection
    Set toRemove = New Collection

    For Each comp In vbProj.VBComponents
        Select Case comp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                ' ※ThisWorkbook/Sheet/フォームは削除しない
                toRemove.Add comp
        End Select
    Next comp

    Dim i As Long
    For i = toRemove.count To 1 Step -1
        vbProj.VBComponents.Remove toRemove(i)
    Next i
End Sub

