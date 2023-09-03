# ライブラリの読み込み
Add-Type -AssemblyName System.Drawing

# 元のファイルがあるディレクトリを指定
$sourceFolder = "N:\cyan\pool"

# 分類されたファイルを格納するディレクトリを指定
$targetFolder = "N:\cyan\pictures"

# メッセージ
Write-Output 写真と動画の整理を開始します
Write-Output ("移動前のディレクトリ: " + $sourceFolder)
Write-Output ("移動後のディレクトリ: " + $targetFolder)

# 画像ファイルと動画ファイルの拡張子を指定
$files = Get-ChildItem -Path $sourceFolder -Recurse | Where-Object { $_.Extension -match "(jpg|jpeg|png|gif|bmp|mp4|mov|avi|wmv|flv|mkv)" }

# 撮影日時を取得する関数
function getTakenDate($file){

    # Exif情報を取得
    $private:properties = $file.PropertyItems
    Write-Output $properties
    foreach($property in $properties){
        if($property.Id -eq 36867){ #0x9003 PropertyTagExifDTOrig
            # $takenDate = $property.Value
            $takenDate = [System.Text.Encoding]::ASCII.GetString($property.Value)
            break
        }
    }

    return $takenDate
}

function getOldest() {
    # WIP
}

function main() {

    # フォルダ内のすべてのファイルに処理を実行
    foreach ($file in $files) {

        # ファイルの作成日時を取得
        $creationDate = $file.CreationTime

        # ファイルの更新日時を取得
        $updateDate = $file.LastWriteTime

        # ファイルの撮影日時を取得
        $takenDate = getTakenDate $file
        Write-Output ("takenDate = " + $takenDate)
        # $takenDate = (Get-ItemProperty $file.FullName).DateTaken
        # $takenDate = (Get-ItemProperty $file.FullName).DateTaken
        # ($img.PropertyItems | Where-Object{$_.Id -eq 36867}).Value

        # 作成日時、更新日時、撮影日時から最も古いものを取得
        $oldestDate = $creationDate
        if ($updateDate -lt $oldestDate) {
            $oldestDate = $updateDate
        }

        # デバッグ
        Write-Output ($file.FullName + "   " + $creationDate + "   " + $updateDate + "   " + $takenDate + "   " + $oldestDate)

        # 年月日を取得
        $year = $creationDate.Year.ToString()
        $month = $creationDate.Month.ToString("00")
        $day = $creationDate.Day.ToString("00")

        # 格納先のディレクトリを作成
        $targetDirectory = $targetFolder + "\" + $year + "\" + $month + "\" + $day

        # ディレクトリが存在しない場合は作成
        if (!(Test-Path $targetDirectory)) {
            New-Item -ItemType Directory -Path $targetDirectory | Out-Null
        }

        # ファイルを移動
        Move-Item -Path $file.FullName -Destination $targetDirectory
        Write-Output ($file.Name + " を移動しました。")
    }
}

# 処理の実行
main
pause
