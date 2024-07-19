
Add-Type -AssemblyName System.Windows.Forms

function Get-ScreenCapture($dir, $name)
{   
    begin {
        Add-Type -AssemblyName System.Drawing, System.Windows.Forms
        $jpegCodec = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | 
            Where-Object { $_.FormatDescription -eq "JPEG" }
    }
    process {
        Start-Sleep -Milliseconds 250

        #Alt+PrintScreenを送信
        [Windows.Forms.Sendkeys]::SendWait("%{PrtSc}")        

        Start-Sleep -Milliseconds 250

        #クリップボードから画像をコピー
        $bitmap = [Windows.Forms.Clipboard]::GetImage()    

        #画像保存(名前がかぶらないようにしている)
        $ep = New-Object Drawing.Imaging.EncoderParameters  
        $ep.Param[0] = New-Object Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality, [long]100)
        $screenCapturePathBase = "$dir\$name"
        $c = 0
        while (Test-Path "${screenCapturePathBase}_${c}.jpg") {
            $c++
        }
        $bitmap.Save("${screenCapturePathBase}_${c}.jpg", $jpegCodec, $ep)
    }
}

$ps = Get-Process msedge
foreach($process in $ps){
    try {
        [Microsoft.VisualBasic.Interaction]::AppActivate($process.Id)
        for ($i = 0; $i -lt $args[1]; $i++) {
            $date = Get-Date -format "yyyyMMddHHmmss"
            Get-ScreenCapture $args[0] $date 
            [System.Windows.Forms.SendKeys]::SendWait("^{TAB}")
        }
    }
    catch {
        
    }
}
