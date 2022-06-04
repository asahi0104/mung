using namespace System.Windows.Forms
Add-Type -Assembly System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function mung([string[]]$matikane_array) {
  # カスタムオブジェクトの配列
  $shiraokis =  @()

  # カスタムオブジェクトクラス
  class CustomObject {
    [String] $TARGET_NAME
    [String] $RESULT
  }
  
  foreach($matikane in $matikane_array) {
    # ホスト名を引っこ抜く
    $tmp = $matikane.Split(" ")
    $tannhauser = $tmp[1]

    # カスタムオブジェクト作成
    $shiraoki = New-Object CustomObject
    $shiraoki.TARGET_NAME = $tannhauser

    $mun_return = Test-Connection $tannhauser -Quiet

    if ($mun_return) {
      $shiraoki.RESULT = "むん！"
    } else {
      $shiraoki.RESULT = "はぐわぁっ！？"
    }

    $shiraokis += $shiraoki
  }
  $shiraokis | Out-GridView
}


# hostsファイル
$mun_hosts_file = "C:\Windows\System32\drivers\etc\hosts"

# ファイルが無かったらエラー
if ( -not (Test-Path $mun_hosts_file )) {
  [System.Windows.Forms.MessageBox]::Show("hostsファイルが見つかりません！$mun_hosts_file", "はぐわぁっ！？","OK","Error","button1") > $null
  exit
}

# hostsファイルの中身取得
$mung_targets = $(Get-Content $mun_hosts_file | Select-String "^#" -NotMatch | Select-String "^\s*$" -NotMatch)

# 中身が空か確認
if([String]::IsNullOrEmpty($mung_targets.Insert)){
  [System.Windows.Forms.MessageBox]::Show("レコードがありません！", "はぐわぁっ！？","OK","Error","button1") > $null
  exit
}

# フォーム作成
$mun_frame = New-Object Form -Property @{
  Text            = 'トマト！'
  Size            = New-Object Drawing.Size(650,400)
  MaximizeBox     = $false
  FormBorderStyle = 'FixedDialog'
  Font            = New-Object Drawing.Font('Meiryo UI', 8.5)
}

# ラベル
$label_mung = New-Object Label -Property @{
  Text     = 'mungを実行したいホストを選んでください！'
  Location = New-Object Drawing.Point(20,20)
  AutoSize = $true
}
$label_mung2 = New-Object Label -Property @{
  Location = New-Object Drawing.Point(20,60)
  AutoSize = $true
}

$mun_frame.Controls.AddRange(@($label_mung,$label_mung2))

# リストボックス定義
$mung_list_box = New-Object ListBox -Property @{
  Location      = New-Object Drawing.Point(20,50)
  Size          = New-Object Drawing.Size(200,300)
  SelectionMode = 'MultiExtended'
  Sorted        = $true
}

foreach($mun_record in $mung_targets) {
  [void]$mung_list_box.Items.Add($mun_record)
}

$mung_target_box = New-Object ListBox -Property @{
  Location      = New-Object Drawing.Point(300,50)
  Size          = New-Object Drawing.Size(200,300)
  SelectionMode = 'MultiExtended'
  Sorted        = $true
}

$mun_frame.Controls.Add($mung_list_box)
$mun_frame.Controls.Add($mung_target_box)

# 追加ボタン
$munn_add = New-Object Button -Property @{
  Text     = '>>'
  Location = New-Object Drawing.Point(230,160)
  Width    = 60
}

# 削除ボタン
$munn_delete = New-Object Button -Property @{
  Text     = '<<'
  Location = New-Object Drawing.Point(230,200)
  Width    = 60
}

# mungボタン
$munn_mung = New-Object Button -Property @{
  Text     = 'むん！'
  Location = New-Object Drawing.Point(515,325)
  Width    = 115
}

# 追加処理
$munn_add.Add_Click({
  for($i = 0; $i -lt $mung_list_box.SelectedItems.Count; $i++) {
    [void]$mung_target_box.Items.Add($mung_list_box.SelectedItems.Item($i))
  }
  while ($mung_list_box.SelectedIndex -ge 0) {
    [void]$mung_list_box.Items.RemoveAt($mung_list_box.SelectedIndex)
  }
})

# 削除処理
$munn_delete.Add_Click({
  for($j = 0; $j -lt $mung_target_box.SelectedItems.Count; $j++) {
    [void]$mung_list_box.Items.Add($mung_target_box.SelectedItems.Item($j))
  }
  while ($mung_target_box.SelectedIndex -ge 0) {
    [void]$mung_target_box.Items.RemoveAt($mung_target_box.SelectedIndex)
  }
})

# mung
$munn_mung.Add_Click({
  $mung_array = @()
  for($k = 0; $k -lt $mung_target_box.Items.Count; $k++) {
    $mung_array += $mung_target_box.Items.Item($k)
  }
  # 何も選ばれてなかったらメッセージを出す
  if ($mung_array.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("対象を選んで下さい！", "なむなむ～！","OK","Information","button1") > $null
  } else {
    $mun_judge = [System.Windows.Forms.MessageBox]::Show("実行します！　良いですか～？", "かくにん！","YesNo","Question","button1")

    if($mun_judge -eq "Yes") {
      mung $mung_array
    }
  }
})

$mun_frame.Controls.Add($munn_add)
$mun_frame.Controls.Add($munn_delete)
$mun_frame.Controls.Add($munn_mung)

$mun_frame.ShowDialog() > $null
