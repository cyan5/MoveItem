# ���̃t�@�C��������t�H���_���w��
$SourceFolder = "C:\Users\cyan\Pictures\test\dummyfolder - �R�s�["

# ���ނ��ꂽ�t�@�C�����i�[����t�H���_���w��
$TargetFolder = "C:\Users\cyan\Pictures\test\sortedfolder"

# ���O�o�͐�
$LogFile = "$TargetFolder\Log.txt"


########################

function Log($LogFile, $str) {
    Write-Output $str
    
    # ���O���o�͂��Ȃ��ꍇ�͈ȉ����R�����g�A�E�g
    Add-Content  $str -Path $LogFile -Encoding UTF8
}

function getTakenDate($SourceImage) {
    Add-Type -AssemblyName System.Drawing

    # �B�e�������o�C�g��Ŏ擾
    $ByteAry = ($SourceImage.PropertyItems | Where-Object{$_.Id -eq 36867}).Value

    # �B�e�������擾�ł����Ƃ�
    if ($null -ne $ByteAry) {

        # ���t�̋�؂蕶����:����/�ɕϊ�
        $ByteAry[4] = 47
        $ByteAry[7] = 47

        # �o�C�g��𕶎���ɕϊ�
        [String]$DateStr = [System.Text.Encoding]::ASCII.GetString($ByteAry)

        # Exif���̓��t��ݒ�
        $TakenDate = [datetime]$DateStr

        return $TakenDate

    } else {

        return $null

    }
}

function main() {

    # �t�H���_�A�T�u�t�H���_���̃t�@�C�������ׂĕ\��
    $Files = Get-ChildItem -Path $SourceFolder -Recurse | Where-Object { $_.Extension -match "(jpg|jpeg|png|gif|bmp|heic|mp4|mov|avi|wmv|flv|mkv)" }

    # �t�H���_�A�T�u�t�H���_���́A�g���q���}�b�`����t�@�C�������ׂĕ\��
    $Pics = $Files | Where-Object { $_.Extension -match "(jpg|jpeg|png|gif|bmp|heic|mp4|mov|avi|wmv|flv|mkv)" }

    # �J�E���^�[�֘A
    $BeforePicsCount = $Pics.Count  # �ʐ^�Ɠ���t�@�C���̍��v��
    $SuccessCount    = 0            # �t�@�C���ړ��ɐ���������

    # ���O
    $Date = (Get-Date -Format "yyyy/MM/dd HH:mm:ss") + " >"
    $Empty = "                     "

    # ���O�o��
    if ($Files.Count -eq 0) {
        Log $LogFile "$Date �t�H���_�͋�ł��B"
    }

    if ($Files.Count - $Pics.Count -gt 0){
        Log $LogFile "$Date " + ($Files.Count - $Pics.Count) + "�̃t�@�C���͎ʐ^�܂��͓���t�@�C���ł͂���܂���B"
    }

    if ($Pics.Count -eq 0) {
        if ($Files.Count -ne 0) {
            Log $LogFile "$Date �ړ�����t�@�C���͂���܂���B"
        }
    } else {
        Log $LogFile "$Date �ʐ^�Ɠ���t�@�C���̈ړ����J�n���܂��B`n$Empty �ړ��O�̃t�H���_: $SourceFolder`n$Empty �ړ���̃t�H���_: $TargetFolder\~"
        
        # �t�H���_���̂��ׂẴt�@�C���ɏ��������s
        $LoopIndex = 0
        foreach ($File in $Pics) {
            
            # �쐬�������擾
            $CreationDate = $File.CreationTime
            
            # �X�V�������擾
            $UpdateDate = $File.LastWriteTime
            
            # �B�e�������擾
            $TakenDate = $null
            if ($File.Extention -match "(jpg|jpeg|heic)") {
                
                $SourceImage = New-Object System.Drawing.Bitmap($File.FullName)
                try {
                    $SourceImage = New-Object System.Drawing.Bitmap($File.FullName)
                } catch {
                    Log $LogFile "$Date $File �̓ǂݍ��݂Ɏ��s���܂����B"
                }
                $TakenDate = getTakenDate $SourceImage
                $SourceImage.Dispose()
            }

            # �ł��Â��������擾
            $OldestDate = $CreationDate
            if ($UpdateDate -lt $OldestDate) {
                $OldestDate = $UpdateDate
            }
            if ($null -ne $TakenDate) {
                if ($TakenDate -lt $OldestDate) {
                    $OldestDate = $TakenDate
                }
            }
            
            # �N�������擾
            $Year  = $OldestDate.Year.ToString()
            $Month = $OldestDate.Month.ToString("00")
            $Day   = $OldestDate.Day.ToString("00")
            
            # �i�[��̃p�X���쐬
            $SubDirectory = "$Year\$Month\$Year-$Month-$Day"
            $TargetDirectory = "$TargetFolder\$SubDirectory"
            
            # �t�H���_�����݂��Ȃ��ꍇ�͍쐬
            if (!(Test-Path $TargetDirectory)) {
                try {
                    New-Item -ItemType Directory -Path $TargetDirectory | Out-Null
                    
                    # Log $LogFile "$Date $TargetDirectory ���쐬���܂����B"
                    
                } catch {
                    Log $LogFile "$Date $TargetDirectory ���쐬�ł��܂���ł����B"
                }
            }
            
            # ���O�ɏo�͂���t�@�C�������쐬
            if ($File.Name.Length -gt 30) {
                # �����������30�ŃJ�b�g���ċ󔒖���
                $PaddedFileName = $File.Name.Remove(30).PadRight(30)
            } else {
                # �󔒖���
                $PaddedFileName = $File.Name.PadRight(30)
            }
            
            # �t�@�C�����ړ�
            try {
                Move-Item -Path $File.FullName -Destination $TargetDirectory
                
                # Log $LogFile "$Date $PaddedFileName �� ~\$SubDirectory �ֈړ�����܂����B"
                $SuccessCount++
            } catch {
                Log $LogFile "$Date $PaddedFileName �� ~\$TargetDirectory �ւ̈ړ��Ɏ��s���܂����B"
            }
            
            # �v���O���X�o�[��\��
            $LoopIndex++
            $Percent = [int32]($LoopIndex*100/$BeforePicsCount)
            Write-Progress -Activity "Move in Progress" -Status "$LoopIndex/$BeforePicsCount Complete." -PercentComplete $Percent
        }
        
        # ���O�o��
        # �t�H���_�A�T�u�t�H���_���̃t�@�C��
        $Files = Get-ChildItem -Path $SourceFolder -Recurse | Where-Object { $_.Extension -match "(jpg|jpeg|png|gif|bmp|heic|mp4|mov|avi|wmv|flv|mkv)" }

        # �t�H���_�A�T�u�t�H���_���́A�g���q���}�b�`����t�@�C��
        $Pics = $Files | Where-Object { $_.Extension -match "(jpg|jpeg|png|gif|bmp|heic|mp4|mov|avi|wmv|flv|mkv)" }

        Log $LogFile "$Date $SuccessCount �̃t�@�C���̈ړ����������܂����B"
        
        if ($Pics.Count -ne 0) {
            Log $LogFile ("$Date " + [String]$Pics.Count + " �̃t�@�C���͈ړ��ł��܂���ł����B")
        }
    }
}

main
pause
