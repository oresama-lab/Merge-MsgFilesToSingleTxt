# Windows Forms アセンブリをロード
Add-Type -AssemblyName System.Windows.Forms

# フォームを作成
$form = New-Object System.Windows.Forms.Form
$form.Text = 'MSG to TXT Converter'
$form.Size = New-Object System.Drawing.Size(400,200)
$form.StartPosition = 'CenterScreen'

# フォルダ選択ボタンを追加
$folderButton = New-Object System.Windows.Forms.Button
$folderButton.Location = New-Object System.Drawing.Point(10,20)
$folderButton.Size = New-Object System.Drawing.Size(120,25)
$folderButton.Text = 'Select Folder'
$folderButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.ShowDialog() | Out-Null
    $textBox1.Text = $folderBrowser.SelectedPath
})
$form.Controls.Add($folderButton)

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(140,20)
$textBox1.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBox1)

# ファイル選択ボタンを追加
$fileButton = New-Object System.Windows.Forms.Button
$fileButton.Location = New-Object System.Drawing.Point(10,60)
$fileButton.Size = New-Object System.Drawing.Size(120,25)
$fileButton.Text = 'Select Output File'
$fileButton.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $fileDialog.Filter = "Text Files (*.txt)|*.txt"
    $fileDialog.ShowDialog() | Out-Null
    $textBox2.Text = $fileDialog.FileName
})
$form.Controls.Add($fileButton)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(140,60)
$textBox2.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBox2)

# 変換ボタンを追加
$convertButton = New-Object System.Windows.Forms.Button
$convertButton.Location = New-Object System.Drawing.Point(140,100)
$convertButton.Size = New-Object System.Drawing.Size(100,30)
$convertButton.Text = 'Convert'
$convertButton.Add_Click({

    # Outlook アプリケーションを開く
    $outlook = New-Object -ComObject Outlook.Application

    # .msg ファイルが存在するフォルダ
    $folderPath = $textBox1.Text

    # 出力する .txt ファイルのパス
    $txtPath = $textBox2.Text

    # フォルダおよびそのサブフォルダ内のすべての .msg ファイルを取得
    $msgFiles = Get-ChildItem -Path $folderPath -Filter "*.msg" -Recurse

    $totalFiles = $msgFiles.Count
    $currentFile = 0

    # 各 .msg ファイルの内容をテキストファイルに追加
    foreach ($msgFile in $msgFiles) {
        $currentFile++

        # 進捗状況を表示
        Write-Progress -Activity "Processing .msg files" -Status "$currentFile of $totalFiles" -PercentComplete (($currentFile / $totalFiles) * 100)

        # 区切りヘッダーを追加
        "###############################################" | Out-File $txtPath -Append
        "# $($msgFile.Name)" | Out-File $txtPath -Append
        "###############################################" | Out-File $txtPath -Append

        # .msg ファイルを開き、内容をテキストファイルに追加
        $message = $outlook.Session.OpenSharedItem($msgFile.FullName)
        $message.Body | Out-File $txtPath -Append

        # Outlook オブジェクトを解放
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($message) | Out-Null
    }

    # 進捗バーをクリア
    Write-Progress -Activity "Processing .msg files" -Completed

    # Outlook オブジェクトを解放
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # ポップアップメッセージを表示
    [System.Windows.Forms.MessageBox]::Show("処理が完了しました！", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})
$form.Controls.Add($convertButton)

# フォームを表示
$form.ShowDialog()
