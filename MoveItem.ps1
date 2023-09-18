# 仕分け前のフォルダを指定
$SourceFolder = "C:\Users\cyan\Pictures\test\dummyfolder"

# 仕分け後のフォルダを指定
$TargetFolder = "C:\Users\cyan\Pictures\test\sortedfolder"

# ログファイルを出力するかどうか ( $true / $false )
$LogFileFlag = $true

# ログファイルの出力パス
$LogFilePath = "$TargetFolder\Log.txt"


######## ここから先はいじらない ########


Add-Type -AssemblyName System.Drawing

function Log($str) {
    $date = (Get-Date -Format "yyyy/MM/dd HH:mm:ss") + " > "
    Write-Host $date$str
        
    # ログファイルを出力
    if ($LogFileFlag) {
        Add-Content $date$str -Path $LogFilePath -Encoding UTF8
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
    if ($SuccessCount -gt 0) {
        Log "$SuccessCount 個のファイルの移動が完了しました。"
    }
    if ($FailureCount -gt 0) {
        Log "$FailureCount 個のファイルは移動できませんでした。"
        pause
    }
}

# Exifから撮影日時を取得する
function GetExifDate($SourceImage) {

    # 撮影日時をバイト列で取得
    $ByteAry = ($SourceImage.PropertyItems | Where-Object {$_.Id -eq 36867}).Value

    # 撮影日時を取得できたとき
    if ($null -ne $ByteAry) {

        # 日付の区切り文字を:から/に変換
        $ByteAry[4] = 47
        $ByteAry[7] = 47

        # バイト列を文字列に変換
        [String]$DateStr = [System.Text.Encoding]::ASCII.GetString($ByteAry)

        # Exif情報の日付を設定
        $TakenDate = [DateTime]$DateStr

        return $TakenDate

    } else {
        return $null
    }
}

# 詳細プロパティから撮影日時を取得する
function GetPropDate($item, $shellObject, $flag) {
    # $flagが0なら撮影日時
    # $flagが0以外ならメディアの作成日時
    if ($flag -eq 0) {
        $PropType = "撮影日時"
    } else {
        $PropType = "メディアの作成日時"
    }

    $folderPath = Split-Path $item.FullName
    $fileName = $item.Name

    $shellFolder = $shellObject.namespace($folderPath)
    $shellFile = $shellFolder.parseName($fileName)
    $selectedPropNo = ""
    $selectedPropValue = ""
    
    # 詳細プロパティの列挙
    for ($i = 0; $i -lt 330; $i++) {
        $PropName = $shellFolder.getDetailsOf($null, $i)
        if ($PropName -eq $PropType) {
            
            $PropValue = $shellFolder.getDetailsOf($shellFile, $i)  # 値の取得
            if ($PropValue) {
                $selectedPropNo = $i
                $selectedPropValue = $PropValue
                break
            }
        }
    }
    if (!$selectedPropNo) {
        return $null
    }

    # " yyyy/ MM/ dd   h:mm" -> "yyyy/MM/dd hh:mm:ss"
    $time = "0" + $selectedPropValue.substring(16) + ":00" # 秒は取得できないので00を設定
    $time = $time.substring($time.length - 8, 8)
    $selectedPropValue = $selectedPropValue.substring(1, 5) + `
                         $selectedPropValue.substring(7, 3) + `
                         $selectedPropValue.substring(11, 2) + " " + $time
    $PropDate = [DateTime]::ParseExact($selectedPropValue, "yyyy/MM/dd HH:mm:ss", $null)

    return $PropDate
}

function main() {

    # フォルダ、サブフォルダ内のファイルをすべて表示
    $Files = Get-ChildItem -Path $SourceFolder -Recurse | Where-Object { ! $_.PsIsContainer }

    # フォルダ、サブフォルダ内の、拡張子がマッチするファイルをすべて表示
    $Items = $Files | Where-Object { $_.Extension -match "(jpg|jpeg|png|gif|bmp|heic|mp4|mov|avi|wmv|flv|mkv)" }

    # カウンター関連
    $BeforeFilesCount = [int32]$Files.Count  # 写真と動画ファイルの合計数
    $BeforeItemsCount = [int32]$Items.Count  # 写真と動画ファイルの合計数
    $SuccessCount     = 0                    # ファイル移動に成功した数
    $FailureCount     = 0                    # ファイル移動に失敗した数

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
            
            # jpg,jpegから撮影日時、それ以外からは撮影日時とメディアの作成日時を取得
            $jpgTakenDate = $null
            $otherTakenDate = $null
            $mediaCreationDate = $null
            if ($item.Extension -match "(jpg|jpeg)") {

                $SourceImage = $null

                try {
                    $SourceImage = New-Object System.Drawing.Bitmap($item.FullName)
                } catch {
                    Log "$item の読み込みに失敗しました。"
                    $FailureCount++
                    continue
                }
                $jpgTakenDate = GetExifDate $SourceImage
                $SourceImage.Dispose()
                $SourceImage = $null

            } else {

                try {
                    $shellObject = New-Object -ComObject Shell.Application
                } catch {
                    Log "$item の読み込みに失敗しました。"
                    $FailureCount++
                    continue
                }

                $otherTakenDate = GetPropDate $item $shellObject 0
                $mediaCreationDate = GetPropDate $item $shellObject 1

                # シェルオブジェクトの解放
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellObject) | out-null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()

            }

            # 最も古い日時を取得
            $OldestDate = $CreationDate
            if ($UpdateDate -lt $OldestDate) {
                $OldestDate = $UpdateDate
            }
            if ($null -ne $jpgTakenDate) {
                if ($jpgTakenDate -lt $OldestDate) {
                    $OldestDate = $jpgTakenDate
                }
            }
            if ($null -ne $otherTakenDate) {
                if ($otherTakenDate -lt $OldestDate) {
                    $OldestDate = $otherTakenDate
                }
            }
            if ($null -ne $mediaCreationDate) {
                if ($mediaCreationDate -lt $OldestDate) {
                    $OldestDate = $mediaCreationDate
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
                    New-Item -ItemType Directory -Path $TargetPath | Out-Null
                    
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
                Move-Item -Path $item.FullName -Destination $TargetPath
                
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
