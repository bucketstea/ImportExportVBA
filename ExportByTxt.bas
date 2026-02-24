Attribute VB_Name = "ExportByTxt"
Option Explicit

Public Sub ExportVbaFilesAsTxt_Modern()
    ' ==========================================
    ' 入力元と出力先のフォルダパスを指定
    ' ==========================================
    Dim srcPath As String: srcPath = ActiveSheet.Cells(11, 2)
    Dim outPath As String: outPath = ActiveSheet.Cells(12, 2)

    ' FileSystemObjectの準備（直前宣言＆即代入）
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    ' ソースフォルダの存在チェック
    If Not fso.FolderExists(srcPath) Then
        MsgBox "ソースフォルダが見つかりません:" & vbCrLf & srcPath, vbExclamation, "エラー"
        Exit Sub
    End If

    ' 出力フォルダがなければ作成
    If Not fso.FolderExists(outPath) Then
        fso.CreateFolder outPath
    End If

    Dim srcFolder As Object: Set srcFolder = fso.GetFolder(srcPath)
    Dim count As Integer: count = 0

    Dim file As Object
    For Each file In srcFolder.Files
        
        ' 拡張子を取得
        Dim ext As String: ext = LCase(fso.GetExtensionName(file.Name))

        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            
            ' 新しいファイル名と出力パスの生成
            Dim newFileName As String: newFileName = fso.GetBaseName(file.Name) & "_" & ext & ".txt"
            Dim destPath As String: destPath = fso.BuildPath(outPath, newFileName)

            ' ファイルのコピーとカウントアップ
            fso.CopyFile file.Path, destPath, True
            count = count + 1
            
        End If
    Next file

    ' 完了メッセージ
    MsgBox count & " 件のファイルをテキスト形式に変換（コピー）しました！", vbInformation, "処理完了"

    ' オブジェクトの解放
    Set fso = Nothing
    Set srcFolder = Nothing
End Sub
