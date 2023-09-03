# ���C�u�����̓ǂݍ���
Add-Type -AssemblyName System.Drawing

# ���̃t�@�C��������f�B���N�g�����w��
$sourceFolder = "N:\cyan\pool"

# ���ނ��ꂽ�t�@�C�����i�[����f�B���N�g�����w��
$targetFolder = "N:\cyan\pictures"

# ���b�Z�[�W
Write-Output �ʐ^�Ɠ���̐������J�n���܂�
Write-Output ("�ړ��O�̃f�B���N�g��: " + $sourceFolder)
Write-Output ("�ړ���̃f�B���N�g��: " + $targetFolder)

# �摜�t�@�C���Ɠ���t�@�C���̊g���q���w��
$files = Get-ChildItem -Path $sourceFolder -Recurse | Where-Object { $_.Extension -match "(jpg|jpeg|png|gif|bmp|mp4|mov|avi|wmv|flv|mkv)" }

# �B�e�������擾����֐�
function getTakenDate($file){

    # Exif�����擾
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

    # �t�H���_���̂��ׂẴt�@�C���ɏ��������s
    foreach ($file in $files) {

        # �t�@�C���̍쐬�������擾
        $creationDate = $file.CreationTime

        # �t�@�C���̍X�V�������擾
        $updateDate = $file.LastWriteTime

        # �t�@�C���̎B�e�������擾
        $takenDate = getTakenDate $file
        Write-Output ("takenDate = " + $takenDate)
        # $takenDate = (Get-ItemProperty $file.FullName).DateTaken
        # $takenDate = (Get-ItemProperty $file.FullName).DateTaken
        # ($img.PropertyItems | Where-Object{$_.Id -eq 36867}).Value

        # �쐬�����A�X�V�����A�B�e��������ł��Â����̂��擾
        $oldestDate = $creationDate
        if ($updateDate -lt $oldestDate) {
            $oldestDate = $updateDate
        }

        # �f�o�b�O
        Write-Output ($file.FullName + "   " + $creationDate + "   " + $updateDate + "   " + $takenDate + "   " + $oldestDate)

        # �N�������擾
        $year = $creationDate.Year.ToString()
        $month = $creationDate.Month.ToString("00")
        $day = $creationDate.Day.ToString("00")

        # �i�[��̃f�B���N�g�����쐬
        $targetDirectory = $targetFolder + "\" + $year + "\" + $month + "\" + $day

        # �f�B���N�g�������݂��Ȃ��ꍇ�͍쐬
        if (!(Test-Path $targetDirectory)) {
            New-Item -ItemType Directory -Path $targetDirectory | Out-Null
        }

        # �t�@�C�����ړ�
        Move-Item -Path $file.FullName -Destination $targetDirectory
        Write-Output ($file.Name + " ���ړ����܂����B")
    }
}

# �����̎��s
main
pause
