# 元のファイルがあるフォルダを指定
$SourceFolder = "C:\Users\cyan\Pictures\test\dummyfolder - コピー"

# 分類されたファイルを格納するフォルダを指定
$TargetFolder = "C:\Users\cyan\Pictures\test\sortedfolder"

# ログファイルの出力パス
$LogFilePath = "$TargetFolder\Log.txt"

# ログファイルを出力するかどうか ( $true / $false )
$LogFileFlag = $true


######## ここから先はいじらない ########


function Log($str) {
    $date = (Get-Date -Format "yyyy/MM/dd HH:mm:ss") + " > "
    Write-Host $date$str
        
    # ログファイルを出力
    if ($LogFileFlag) {
        Add-Content $str -Path $LogFilePath -Encoding UTF8
    }
}

function MessageBefore($FilesCount, $ItemsCount) {

    if ($FilesCount -eq 0) {    # フォルダが空のとき
        Log "フォルダは空です。"
    } else {    # フォルダが空ではないとき

        # 写真または動画ではないファイル
        $OtherFileCount = $FilesCount - $ItemsCount
        if ($OtherFileCount -gt 0) {
            Log "$OtherFileCount 個のファイルは写真または動画ファイルではありません。"
        }

        # 写真または動画
        if ($ItemsCount -eq 0) {
            Log "移動するファイルはありません。"
        } else {
            Log ([String]$ItemsCount + " 個の写真または動画ファイルが見つかりました。ファイルの移動を開始します。")
        }

    }
}

function MessageAfter($SuccessCount, $FailureCount) {
    Log "$SuccessCount 個のファイルの移動が完了しました。"
    if ($FailureCount -ne 0) {
        Log "$FailureCount 個のファイルは移動できませんでした。"
    }
}

function GetTakenDate($SourceImage) {
    Add-Type -AssemblyName System.Drawing

    # 撮影日時をバイト列で取得
    $ByteAry = ($SourceImage.PropertyItems | Where-Object{$_.Id -eq 36867}).Value

    # 撮影日時を取得できたとき
    if ($null -ne $ByteAry) {

        # 日付の区切り文字を:から/に変換
        $ByteAry[4] = 47
        $ByteAry[7] = 47

        # バイト列を文字列に変換
        [String]$DateStr = [System.Text.Encoding]::ASCII.GetString($ByteAry)

        # Exif情報の日付を設定
        $TakenDate = [datetime]$DateStr

        return $TakenDate

    } else {

        return $null

    }
}

function main() {

    # フォルダ、サブフォルダ内のファイルをすべて表示
    $Files = Get-ChildItem -Path $SourceFolder -Recurse | Where-Object { $_.Extension -match "(jpg|jpeg|png|gif|bmp|heic|mp4|mov|avi|wmv|flv|mkv)" }

    # フォルダ、サブフォルダ内の、拡張子がマッチするファイルをすべて表示
    $Items = $Files | Where-Object { $_.Extension -match "(jpg|jpeg|png|gif|bmp|heic|mp4|mov|avi|wmv|flv|mkv)" }

    # カウンター関連
    $BeforeFilesCount = [int32]$Files.Count  # 写真と動画ファイルの合計数
    $BeforeItemsCount = [int32]$Items.Count  # 写真と動画ファイルの合計数
    $SuccessCount     = 0                    # ファイル移動に成功した数
    $FailureCount       = 0                    # ファイル移動に失敗した数

    # ログ出力
    MessageBefore $BeforeFilesCount $BeforeItemsCount

    # 以下、写真または動画ファイルがあるときのみ実行
    if ($BeforeItemsCount -gt 0) {
        
        # フォルダ内のすべてのファイルに処理を実行
        $LoopIndex = 0
        foreach ($item in $Items) {

            # プログレスバーを表示
            $Percent = [int32]($LoopIndex*100/$BeforeItemsCount)
            Write-Progress -Activity "Move in Progress" -Status "$LoopIndex/$BeforeItemsCount Complete." -PercentComplete $Percent
            
            # 作成日時を取得
            $CreationDate = $item.CreationTime
            
            # 更新日時を取得
            $UpdateDate = $item.LastWriteTime
            
            # 撮影日時を取得
            $TakenDate = $null
            if ($item.Extention -match "(jpg|jpeg|heic)") {
                
                $SourceImage = New-Object System.Drawing.Bitmap($item.FullName)
                try {
                    $SourceImage = New-Object System.Drawing.Bitmap($item.FullName)
                } catch {
                    Log "$item の読み込みに失敗しました。"
                }
                $TakenDate = GetTakenDate $SourceImage
                $SourceImage.Dispose()
            }

            # 最も古い日時を取得
            $OldestDate = $CreationDate
            if ($UpdateDate -lt $OldestDate) {
                $OldestDate = $UpdateDate
            }
            if ($null -ne $TakenDate) {
                if ($TakenDate -lt $OldestDate) {
                    $OldestDate = $TakenDate
                }
            }
            
            # 年月日を取得
            $Year  = $OldestDate.Year.ToString()
            $Month = $OldestDate.Month.ToString("00")
            $Day   = $OldestDate.Day.ToString("00")
            
            # 格納先のパスを作成
            $SubPath = "$Year\$Month\$Year-$Month-$Day"
            $TargetPath = "$TargetFolder\$SubPath"
            
            # フォルダが存在しない場合は作成
            if (!(Test-Path $TargetPath)) {
                try {
                    # New-Item -ItemType Directory -Path $TargetPath | Out-Null
                    
                    # Log "$TargetPath を作成しました。"
                    
                } catch {
                    Log "$TargetPath を作成できませんでした。"
                }
            }
            
            # ログに出力するファイル名を作成
            if ($item.Name.Length -gt 30) {
                $PaddedFileName = $item.Name.Remove(30).PadRight(30)   # 長い文字列を30でカットして空白埋め
            } else {
                $PaddedFileName = $item.Name.PadRight(30)              # 空白埋め
            }
            
            # ファイルを移動
            try {
                # Move-Item -Path $item.FullName -Destination $TargetPath
                
                # Log "$PaddedFileName は ~\$SubPath へ移動されました。"
                $SuccessCount++
            } catch {
                Log "$PaddedFileName は ~\$TargetPath への移動に失敗しました。"
                $FailureCount++
            }
            
            $LoopIndex++
        }
        
        # ログ出力
        MessageAfter $SuccessCount $FailureCount
    }
}

main
pause
