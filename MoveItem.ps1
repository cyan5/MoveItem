# 元のファイルがあるフォルダを指定
$SourceFolder = "C:\Users\cyan\Pictures\test\dummyfolder - コピー"

# 分類されたファイルを格納するフォルダを指定
$TargetFolder = "C:\Users\cyan\Pictures\test\sortedfolder"

# ログ出力先
$LogFile = "$TargetFolder\Log.txt"


########################

function Log($LogFile, $str) {
    Write-Output $str
    
    # ログを出力しない場合は以下をコメントアウト
    Add-Content  $str -Path $LogFile -Encoding UTF8
}

function getTakenDate($SourceImage) {
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
    $Pics = $Files | Where-Object { $_.Extension -match "(jpg|jpeg|png|gif|bmp|heic|mp4|mov|avi|wmv|flv|mkv)" }

    # カウンター関連
    $BeforePicsCount = $Pics.Count  # 写真と動画ファイルの合計数
    $SuccessCount    = 0            # ファイル移動に成功した数

    # ログ
    $Date = (Get-Date -Format "yyyy/MM/dd HH:mm:ss") + " >"
    $Empty = "                     "

    # ログ出力
    if ($Files.Count -eq 0) {
        Log $LogFile "$Date フォルダは空です。"
    }

    if ($Files.Count - $Pics.Count -gt 0){
        Log $LogFile "$Date " + ($Files.Count - $Pics.Count) + "個のファイルは写真または動画ファイルではありません。"
    }

    if ($Pics.Count -eq 0) {
        if ($Files.Count -ne 0) {
            Log $LogFile "$Date 移動するファイルはありません。"
        }
    } else {
        Log $LogFile "$Date 写真と動画ファイルの移動を開始します。`n$Empty 移動前のフォルダ: $SourceFolder`n$Empty 移動先のフォルダ: $TargetFolder\~"
        
        # フォルダ内のすべてのファイルに処理を実行
        $LoopIndex = 0
        foreach ($File in $Pics) {
            
            # 作成日時を取得
            $CreationDate = $File.CreationTime
            
            # 更新日時を取得
            $UpdateDate = $File.LastWriteTime
            
            # 撮影日時を取得
            $TakenDate = $null
            if ($File.Extention -match "(jpg|jpeg|heic)") {
                
                $SourceImage = New-Object System.Drawing.Bitmap($File.FullName)
                try {
                    $SourceImage = New-Object System.Drawing.Bitmap($File.FullName)
                } catch {
                    Log $LogFile "$Date $File の読み込みに失敗しました。"
                }
                $TakenDate = getTakenDate $SourceImage
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
            $SubDirectory = "$Year\$Month\$Year-$Month-$Day"
            $TargetDirectory = "$TargetFolder\$SubDirectory"
            
            # フォルダが存在しない場合は作成
            if (!(Test-Path $TargetDirectory)) {
                try {
                    New-Item -ItemType Directory -Path $TargetDirectory | Out-Null
                    
                    # Log $LogFile "$Date $TargetDirectory を作成しました。"
                    
                } catch {
                    Log $LogFile "$Date $TargetDirectory を作成できませんでした。"
                }
            }
            
            # ログに出力するファイル名を作成
            if ($File.Name.Length -gt 30) {
                # 長い文字列を30でカットして空白埋め
                $PaddedFileName = $File.Name.Remove(30).PadRight(30)
            } else {
                # 空白埋め
                $PaddedFileName = $File.Name.PadRight(30)
            }
            
            # ファイルを移動
            try {
                Move-Item -Path $File.FullName -Destination $TargetDirectory
                
                # Log $LogFile "$Date $PaddedFileName は ~\$SubDirectory へ移動されました。"
                $SuccessCount++
            } catch {
                Log $LogFile "$Date $PaddedFileName は ~\$TargetDirectory への移動に失敗しました。"
            }
            
            # プログレスバーを表示
            $LoopIndex++
            $Percent = [int32]($LoopIndex*100/$BeforePicsCount)
            Write-Progress -Activity "Move in Progress" -Status "$LoopIndex/$BeforePicsCount Complete." -PercentComplete $Percent
        }
        
        # ログ出力
        # フォルダ、サブフォルダ内のファイル
        $Files = Get-ChildItem -Path $SourceFolder -Recurse | Where-Object { $_.Extension -match "(jpg|jpeg|png|gif|bmp|heic|mp4|mov|avi|wmv|flv|mkv)" }

        # フォルダ、サブフォルダ内の、拡張子がマッチするファイル
        $Pics = $Files | Where-Object { $_.Extension -match "(jpg|jpeg|png|gif|bmp|heic|mp4|mov|avi|wmv|flv|mkv)" }

        Log $LogFile "$Date $SuccessCount 個のファイルの移動が完了しました。"
        
        if ($Pics.Count -ne 0) {
            Log $LogFile ("$Date " + [String]$Pics.Count + " 個のファイルは移動できませんでした。")
        }
    }
}

main
pause
