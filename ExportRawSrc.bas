Attribute VB_Name = "ExportRawSrc"
Option Explicit

Public Sub ExportAllModules()
    Dim targetPath As String: targetPath = ActiveSheet.Cells(2, 2)
    
    ' 対象ブックを開く（安全のため読み取り専用で）
    Dim targetBook As Workbook: Set targetBook = Workbooks.Open(Filename:=targetPath, ReadOnly:=True)
    
    ' 出力先ディレクトリ（対象ブックと同じ階層に "_vba_export" フォルダを作成）
    Dim outDir As String: outDir = targetBook.Path & "\_vba_export\"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir
    
    ' 対象ブックのVBProjectを取得
    Dim vbProj As VBIDE.VBProject: Set vbProj = targetBook.VBProject
    
    Dim comp As VBIDE.VBComponent
    For Each comp In vbProj.VBComponents
        ' 特定のモジュールをスキップ_適宜変更して
        If InStr(comp.Name, "Sheet") > 0 Then GoTo Continue
        If InStr(comp.Name, "ThisWorkbook") > 0 Then GoTo Continue
        
        Select Case comp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm, vbext_ct_Document
                comp.Export outDir & comp.Name & ModuleExt(comp.Type)
        End Select
Continue:
    Next comp
    
    ' 対象ブックを閉じる（保存しない）
    targetBook.Close SaveChanges:=False
    
    MsgBox "以下のフォルダにソースを出力しました。: " & vbCrLf & outDir, vbInformation, "エクスポート完了"
End Sub

Private Function ModuleExt(t As VBIDE.vbext_ComponentType) As String
    Select Case t
        Case vbext_ct_StdModule:  ModuleExt = ".bas"
        Case vbext_ct_ClassModule: ModuleExt = ".cls"
        Case vbext_ct_MSForm:     ModuleExt = ".frm" ' Export時に .frx も同時に出力されます
        Case vbext_ct_Document:   ModuleExt = ".cls" ' ThisWorkbook/Sheet
        Case Else:                ModuleExt = ".txt"
    End Select
End Function
