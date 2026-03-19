Attribute VB_Name = "ExportByTxt"
Option Explicit

Public Sub ExportVbaFilesAsTxt_Modern()
    ' ==========================================
    ' 入力元と出力先のフォルダパスを指定
    ' ==========================================
    Dim srcPath As String: srcPath = ActiveSheet.Cells(11, 2).Value
    Dim outPath As String: outPath = ActiveSheet.Cells(12, 2).Value

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    ' フォルダチェック
    If Not fso.FolderExists(srcPath) Then
        MsgBox "ソースフォルダが見つかりません:" & vbCrLf & srcPath, vbExclamation, "エラー"
        Exit Sub
    End If
    If Not fso.FolderExists(outPath) Then fso.CreateFolder outPath

    Dim srcFolder As Object: Set srcFolder = fso.GetFolder(srcPath)
    Dim count As Integer: count = 0
    Dim file As Object

    ' ADODB.Streamオブジェクトの準備
    Dim readStream As Object: Set readStream = CreateObject("ADODB.Stream")
    Dim writeStream As Object: Set writeStream = CreateObject("ADODB.Stream")

    For Each file In srcFolder.Files
        Dim ext As String: ext = LCase(fso.GetExtensionName(file.Name))

        ' 対象拡張子の判定
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            
            ' 出力ファイル名の作成
            Dim newFileName As String: newFileName = fso.GetBaseName(file.Name) & "_" & ext & ".txt"
            Dim destPath As String: destPath = fso.BuildPath(outPath, newFileName)

            ' --- 文字コード変換処理 ---
            
            ' 1. 元ファイルを読み込む (Shift-JISとして読み込み)
            readStream.Open
            readStream.Type = 2 ' adTypeText
            readStream.Charset = "Shift-JIS"
            readStream.LoadFromFile file.Path
            Dim content As String: content = readStream.ReadText
            readStream.Close

            ' 2. UTF-8(BOM有)で書き出す
            writeStream.Open
            writeStream.Type = 2 ' adTypeText
            writeStream.Charset = "UTF-8" ' ADODBはデフォルトでBOMを付与します
            writeStream.WriteText content
            writeStream.SaveToFile destPath, 2 ' adSaveCreateOverWrite
            writeStream.Close
            
            count = count + 1
        End If
    Next file

    MsgBox count & " 件のファイルをBOM付UTF-8形式で保存しました！", vbInformation, "処理完了"

    Set fso = Nothing
    Set readStream = Nothing
    Set writeStream = Nothing
End Sub
