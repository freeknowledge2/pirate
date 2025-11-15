param(
    [switch]$SkipRelaunch
)

try {
    if (-not ([System.Management.Automation.PSTypeName]'DpiUtilities').Type) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class DpiUtilities {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
"@
    }
    [DpiUtilities]::SetProcessDPIAware() | Out-Null
} catch {}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# Relaunch script with ExecutionPolicy Bypass if needed
if (-not $SkipRelaunch -and $PSCommandPath) {
    $workingDirectory = Split-Path -Path $PSCommandPath -Parent
    $argumentList = @(
        '-NoProfile'
        '-ExecutionPolicy','Bypass'
        '-WindowStyle','Hidden'
        '-File', $PSCommandPath
        '-SkipRelaunch'
    )

    Start-Process -FilePath 'powershell.exe' -ArgumentList $argumentList -WorkingDirectory $workingDirectory -WindowStyle Hidden | Out-Null
    exit
}

# Ukrycie konsoli PowerShell
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Window {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
"@

$consolePtr = [Window]::GetConsoleWindow()

# Hook na mysz
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class MouseHook {
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT lpPoint);
    
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT {
        public int X;
        public int Y;
    }
}
"@

# Hook na klawiaturę dla ESC
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class KeyboardHook {
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);
}
"@

# Pobranie rzeczywistej rozdzielczości
$screenBounds = [System.Windows.Forms.SystemInformation]::VirtualScreen


function Get-MonitorInfos {
    $nativeWidth = $null
    $nativeHeight = $null

    try {
        $controllerResult = Get-CimInstance -Namespace root\cimv2 -ClassName Win32_VideoController -ErrorAction Stop
        $controllers = @($controllerResult | Where-Object { $_.CurrentHorizontalResolution -gt 0 -and $_.CurrentVerticalResolution -gt 0 })
        if ($controllers.Count -gt 0) {
            $primaryController = $controllers |
                Sort-Object -Property @{ Expression = { ([int]$_.CurrentHorizontalResolution) * ([int]$_.CurrentVerticalResolution) } } -Descending |
                Select-Object -First 1

            if ($primaryController) {
                $nativeWidth = [int]$primaryController.CurrentHorizontalResolution
                $nativeHeight = [int]$primaryController.CurrentVerticalResolution
            }
        }
    } catch {}

    if (-not $nativeWidth -or $nativeWidth -le 0 -or -not $nativeHeight -or $nativeHeight -le 0) {
        $fallback = [System.Windows.Forms.SystemInformation]::VirtualScreen
        $nativeWidth = [int][Math]::Max(1, $fallback.Width)
        $nativeHeight = [int][Math]::Max(1, $fallback.Height)
    }

    $bounds = New-Object System.Drawing.Rectangle(0, 0, $nativeWidth, $nativeHeight)

    return @([pscustomobject]@{
        Bounds = $bounds
        NativeWidth = [double]$nativeWidth
        NativeHeight = [double]$nativeHeight
        LogicalScaleX = 1.0
        LogicalScaleY = 1.0
        ScaleX = 1.0
        ScaleY = 1.0
        DeviceName = 'Primary'
        ActualLeft = 0.0
        ActualTop = 0.0
        ActualWidth = [double]$nativeWidth
        ActualHeight = [double]$nativeHeight
    })
}

$script:monitorInfos = Get-MonitorInfos
if (-not $script:monitorInfos -or $script:monitorInfos.Count -eq 0) {
    $script:monitorInfos = @([pscustomobject]@{
        Bounds = New-Object System.Drawing.Rectangle($screenBounds.Left, $screenBounds.Top, $screenBounds.Width, $screenBounds.Height)
        NativeWidth = [double]$screenBounds.Width
        NativeHeight = [double]$screenBounds.Height
        LogicalScaleX = 1.0
        LogicalScaleY = 1.0
        ScaleX = 1.0
        ScaleY = 1.0
        DeviceName = 'VirtualScreen'
        ActualLeft = [double]$screenBounds.Left
        ActualTop = [double]$screenBounds.Top
        ActualWidth = [double]$screenBounds.Width
        ActualHeight = [double]$screenBounds.Height
    })
}

if ($script:monitorInfos.Count -gt 0) {
    $firstMonitor = $script:monitorInfos[0]
    $screenBounds = New-Object System.Drawing.Rectangle([int][Math]::Floor($firstMonitor.ActualLeft), [int][Math]::Floor($firstMonitor.ActualTop), [int][Math]::Round($firstMonitor.ActualWidth), [int][Math]::Round($firstMonitor.ActualHeight))
}


[Window]::ShowWindow($consolePtr, 0) | Out-Null

# Ustawienie formularza jako overlay
$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = 'None'
$form.StartPosition = 'Manual'
$form.Location = New-Object System.Drawing.Point($screenBounds.Left, $screenBounds.Top)
$form.Size = New-Object System.Drawing.Size($screenBounds.Width, $screenBounds.Height)
$form.TopMost = $true
$form.BackColor = [System.Drawing.Color]::Magenta
$form.TransparencyKey = [System.Drawing.Color]::Magenta
$form.ShowInTaskbar = $false
$form.Opacity = 1.0

# Ustawienie formularza jako "click-through"
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
}
"@

$GWL_EXSTYLE = -20
$WS_EX_LAYERED = 0x80000
$WS_EX_TRANSPARENT = 0x20

$currentStyle = [Win32]::GetWindowLong($form.Handle, $GWL_EXSTYLE)
[Win32]::SetWindowLong($form.Handle, $GWL_EXSTYLE, $currentStyle -bor $WS_EX_LAYERED -bor $WS_EX_TRANSPARENT) | Out-Null

# PictureBox do wyświetlania efektów
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Location = New-Object System.Drawing.Point(0, 0)
$pictureBox.Size = New-Object System.Drawing.Size($screenBounds.Width, $screenBounds.Height)
$pictureBox.BackColor = [System.Drawing.Color]::Transparent
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Normal
$form.Controls.Add($pictureBox)

$random = New-Object System.Random

# Configurable timing constants (in seconds)
$MIN_DELAY_BETWEEN_GLITCHES = 60
$MAX_DELAY_BETWEEN_GLITCHES = 300

# Lista wszystkich kopii okien do wyświetlenia
$script:windowCopies = New-Object System.Collections.Generic.List[object]
$script:currentCopyIndex = 0
$script:overlayBitmap = $null
$script:exitAnimationStarted = $false
$script:exitAnimationInitializing = $false
$script:exitAnimationType = 0
$script:exitAnimationStep = 0
$script:exitAnimationTimer = [DateTime]::MinValue
$script:nextGlitchTime = [DateTime]::MinValue
$script:waitingForNextGlitch = $false
$script:bsodGlitchBitmaps = New-Object 'System.Collections.Generic.List[System.Drawing.Bitmap]'
$script:pendingExitAnimation = $false
$script:coverCompleteAt = [DateTime]::MinValue
$script:cursorHidden = $false
$script:bsodRevealIndex = 0
$script:bsodRevealTotal = 0
$script:bsodRevealCompleted = $false
$script:bsodSecondStageStartIndex = 0
$script:bsodBeepProcess = $null
$script:plannedExitAnimationType = 1
$script:bsodBeepScheduled = $false
$script:glitchVectors = @(
    [System.Drawing.Point]::new(10, 0),
    [System.Drawing.Point]::new(-10, 0),
    [System.Drawing.Point]::new(0, 10),
    [System.Drawing.Point]::new(0, -10),
    [System.Drawing.Point]::new(10, 10),
    [System.Drawing.Point]::new(-10, 10),
    [System.Drawing.Point]::new(10, -10),
    [System.Drawing.Point]::new(-10, -10),
    [System.Drawing.Point]::new(20, 10),
    [System.Drawing.Point]::new(-20, 10),
    [System.Drawing.Point]::new(20, -10),
    [System.Drawing.Point]::new(-20, -10),
    [System.Drawing.Point]::new(10, 20),
    [System.Drawing.Point]::new(-10, 20),
    [System.Drawing.Point]::new(10, -20),
    [System.Drawing.Point]::new(-10, -20)
)

# Timing controls for the BSOD tone
$BSOD_BEEP_LEAD_SECONDS = 4.0
$BSOD_BEEP_DURATION_SECONDS = 32.0

# Set timer interval for smooth animation (60 FPS)
$timerInterval = 16  # ~60 FPS (1000ms / 60 = 16.67ms)

# Timer do animacji - dopasowany do refresh rate ekranu
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = $timerInterval
$script:lastWindowCapture = [DateTime]::MinValue

# Function to apply filter to bitmap - FAST HEURISTIC VERSION
function Invoke-GlitchFilter {
    param(
        [System.Drawing.Bitmap]$sourceBitmap,
        [int]$filterType
    )
    
    $width = $sourceBitmap.Width
    $height = $sourceBitmap.Height
    $filteredBitmap = New-Object System.Drawing.Bitmap($width, $height, [System.Drawing.Imaging.PixelFormat]::Format32bppPArgb)
    
    # Lock bits for faster pixel access
    $sourceRect = New-Object System.Drawing.Rectangle(0, 0, $width, $height)
    $sourceData = $sourceBitmap.LockBits($sourceRect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $destData = $filteredBitmap.LockBits($sourceRect, [System.Drawing.Imaging.ImageLockMode]::WriteOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppPArgb)
    
    $bytes = [Math]::Abs($sourceData.Stride) * $height
    $sourceBuffer = New-Object byte[] $bytes
    $destBuffer = New-Object byte[] $bytes
    
    [System.Runtime.InteropServices.Marshal]::Copy($sourceData.Scan0, $sourceBuffer, 0, $bytes)
    
    # HEURISTIC: Process every 4 pixels in a block, apply to whole block (4x faster)
    $threshold1 = 150
    $threshold2 = 80
    $threshold3_brightness = 50
    $threshold3_color = 40
    $threshold4 = 0
    
    switch ($filterType) {
        1 { $threshold = $threshold1 }
        2 { $threshold = $threshold2 }
        3 { $threshold = $threshold3_brightness }
        4 { $threshold = $threshold4 }
    }
    
    # Process pixels in buffer (BGRA format) with sampling
    for ($i = 0; $i -lt $bytes; $i += 16) {  # Jump by 4 pixels at a time
        # Sample first pixel of the block
        $b = $sourceBuffer[$i]
        $g = $sourceBuffer[$i + 1]
        $r = $sourceBuffer[$i + 2]
        
        $brightness = ($r + $g + $b) / 3
        $keepPixel = $false
        
        if ($filterType -eq 3) {
            $colorDiff = [Math]::Max([Math]::Abs($r - $g), [Math]::Max([Math]::Abs($g - $b), [Math]::Abs($r - $b)))
            $keepPixel = ($brightness -gt $threshold3_brightness) -or ($colorDiff -gt $threshold3_color)
        } else {
            $keepPixel = $brightness -gt $threshold
        }
        
        # Apply decision to all 4 pixels in the block
        for ($offset = 0; $offset -lt 16; $offset += 4) {
            $idx = $i + $offset
            if ($idx -ge $bytes) { break }
            
            if ($keepPixel) {
                $destBuffer[$idx] = $sourceBuffer[$idx]
                $destBuffer[$idx + 1] = $sourceBuffer[$idx + 1]
                $destBuffer[$idx + 2] = $sourceBuffer[$idx + 2]
                $destBuffer[$idx + 3] = $sourceBuffer[$idx + 3]
            }
        }
    }
    
    [System.Runtime.InteropServices.Marshal]::Copy($destBuffer, 0, $destData.Scan0, $bytes)
    
    $sourceBitmap.UnlockBits($sourceData)
    $filteredBitmap.UnlockBits($destData)
    
    return $filteredBitmap
}

    # Helper: draw single copy onto overlay bitmap
    function Add-GlitchCopy {
        param(
            [System.Drawing.Bitmap]$overlay,
            [hashtable]$copy
        )

        if (-not $overlay -or -not $copy -or -not $copy.Bitmap -or -not $copy.Visible) {
            return
        }

        $graphics = [System.Drawing.Graphics]::FromImage($overlay)
        $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceOver
        $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighSpeed
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::NearestNeighbor
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighSpeed
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::None

        $colorMatrix = New-Object System.Drawing.Imaging.ColorMatrix
        $colorMatrix.Matrix33 = $copy.Opacity
        $imageAttributes = New-Object System.Drawing.Imaging.ImageAttributes
        $imageAttributes.SetColorMatrix($colorMatrix, [System.Drawing.Imaging.ColorMatrixFlag]::Default, [System.Drawing.Imaging.ColorAdjustType]::Bitmap)

        $destRect = New-Object System.Drawing.Rectangle([int]$copy.X, [int]$copy.Y, $copy.Width, $copy.Height)

        try {
            $graphics.DrawImage(
                $copy.Bitmap,
                $destRect,
                0, 0, $copy.Width, $copy.Height,
                [System.Drawing.GraphicsUnit]::Pixel,
                $imageAttributes
            )
        } catch {}

        $imageAttributes.Dispose()
        $graphics.Dispose()
    }

    # Helper: rebuild overlay from all visible copies
    function Update-OverlayComposite {
        param(
            [System.Drawing.Bitmap]$overlay,
            [System.Collections.Generic.List[object]]$copies,
            [System.Drawing.Bitmap]$baseBitmap = $null
        )

        if (-not $overlay) { return }

        $graphics = [System.Drawing.Graphics]::FromImage($overlay)
        if ($baseBitmap) {
            $graphics.DrawImage($baseBitmap, 0, 0)
        } else {
            $graphics.Clear([System.Drawing.Color]::Transparent)
        }
        $graphics.Dispose()

        for ($i = 0; $i -lt $copies.Count; $i++) {
            if ($copies[$i].Visible) {
                Add-GlitchCopy -overlay $overlay -copy $copies[$i]
            }
        }
    }

    function Clear-WindowCopies {
        if ($script:windowCopies -and $script:windowCopies.Count -gt 0) {
            for ($i = 0; $i -lt $script:windowCopies.Count; $i++) {
                $copy = $script:windowCopies[$i]
                if ($copy -and $copy.ContainsKey("OwnsBitmap") -and $copy.OwnsBitmap -and $copy.Bitmap) {
                    try {
                        $copy.Bitmap.Dispose()
                    } catch {}
                }
            }
            $script:windowCopies.Clear()
        }
        $script:currentCopyIndex = 0
    }

    function Reset-Overlay {
        if ($script:overlayBitmap) {
            $pictureBox.Image = $null
            try {
                $script:overlayBitmap.Dispose()
            } catch {}
            $script:overlayBitmap = $null
        }
    }

    function Set-NextGlitchSchedule {
        $script:waitingForNextGlitch = $true
        $delaySeconds = $random.Next($MIN_DELAY_BETWEEN_GLITCHES, $MAX_DELAY_BETWEEN_GLITCHES + 1)
        $script:nextGlitchTime = [DateTime]::Now.AddSeconds($delaySeconds)
        $script:exitAnimationStarted = $false
        $script:exitAnimationInitializing = $false
        $script:exitAnimationStep = 0
        $script:pendingExitAnimation = $false
        $script:coverCompleteAt = [DateTime]::MinValue
        $script:exitAnimationType = 0
        $script:currentCopyIndex = 0
        $script:bsodRevealIndex = 0
        $script:bsodRevealTotal = 0
        $script:bsodRevealCompleted = $false
        $script:bsodSecondStageStartIndex = 0
        Stop-BsodBeep
        if ($script:cursorHidden) {
            [System.Windows.Forms.Cursor]::Show()
            $script:cursorHidden = $false
        }
    }

    function Start-BsodBeep {
        param([double]$DurationSeconds = $BSOD_BEEP_DURATION_SECONDS)

        Stop-BsodBeep
        try {
            $durationMs = [int][Math]::Max(0, [Math]::Round([double]$DurationSeconds * 1000.0))
            $beepCommand = "[console]::Beep(60,{0})" -f $durationMs
            $beepArgs = @(
                '-NoProfile'
                '-Command'
                $beepCommand
            )
            $script:bsodBeepProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList $beepArgs -WindowStyle Hidden -PassThru
        } catch {
            if ($script:bsodBeepProcess) {
                try {
                    if (-not $script:bsodBeepProcess.HasExited) {
                        $script:bsodBeepProcess.Kill()
                    }
                } catch {}
                try { $script:bsodBeepProcess.Dispose() } catch {}
                $script:bsodBeepProcess = $null
            }
        }
    }

    function Clear-BsodGlitches {
        param([bool]$PreserveBeep = $false)
        $script:bsodBeepScheduled = $false
        if ($script:bsodGlitchBitmaps) {
            foreach ($bmp in $script:bsodGlitchBitmaps) {
                if ($bmp) {
                    try {
                        $bmp.Dispose()
                    } catch {}
                }
            }
            $script:bsodGlitchBitmaps.Clear()
        }
        if ($script:bsodBitmap) {
            try {
                $script:bsodBitmap.Dispose()
            } catch {}
            $script:bsodBitmap = $null
        }
        $script:bsodRevealIndex = 0
        $script:bsodRevealTotal = 0
        $script:bsodRevealCompleted = $false
        $script:bsodSecondStageStartIndex = 0
        if (-not $PreserveBeep) {
            Stop-BsodBeep
        }
    }

    function Stop-BsodBeep {
        if ($script:bsodBeepProcess) {
            try {
                if (-not $script:bsodBeepProcess.HasExited) {
                    $script:bsodBeepProcess.Kill()
                }
            } catch {}
            try { $script:bsodBeepProcess.Dispose() } catch {}
            $script:bsodBeepProcess = $null
        }
    }

    function Get-RandomGlitchVector {
        if (-not $script:glitchVectors -or $script:glitchVectors.Count -eq 0) {
            return [System.Drawing.Point]::new(10, 0)
        }
        $index = $random.Next(0, $script:glitchVectors.Count)
        return $script:glitchVectors[$index]
    }

    function Clear-OverlayRegion {
        param(
            [System.Drawing.Rectangle]$region
        )

        if (-not $script:overlayBitmap) { return }

        $graphics = [System.Drawing.Graphics]::FromImage($script:overlayBitmap)
        $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
        $clearBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 0, 0, 0))
        $graphics.FillRectangle($clearBrush, $region)
        $clearBrush.Dispose()
        $graphics.Dispose()
        $pictureBox.Invalidate()
    }

    function New-BsodSectionBitmap {
        param(
            [int]$targetWidth,
            [int]$targetHeight,
            [double]$scaleX,
            [double]$scaleY
        )

        $width = [int][Math]::Max(1, $targetWidth)
        $height = [int][Math]::Max(1, $targetHeight)
        $dpiScaleX = if ($scaleX -gt 0) { $scaleX } else { 1.0 }
        $dpiScaleY = if ($scaleY -gt 0) { $scaleY } else { 1.0 }

        $layoutWidth = [double]$width / $dpiScaleX
        $layoutHeight = [double]$height / $dpiScaleY

        $backgroundColor = [Windows.Media.Color]::FromRgb(0x00, 0x78, 0xD7)
        $baseWidth = 1920.0
        $baseHeight = 1080.0

        $viewbox = New-Object System.Windows.Controls.Viewbox
        $viewbox.Width = $layoutWidth
        $viewbox.Height = $layoutHeight
        $viewbox.Stretch = 'Fill'
        $viewbox.StretchDirection = 'Both'
        $viewbox.SnapsToDevicePixels = $true

        $root = New-Object System.Windows.Controls.Grid
        $root.Width = $baseWidth
        $root.Height = $baseHeight
        $root.Background = New-Object Windows.Media.SolidColorBrush $backgroundColor
        $root.UseLayoutRounding = $true
        [Windows.Media.RenderOptions]::SetBitmapScalingMode($root, [Windows.Media.BitmapScalingMode]::HighQuality)
        [Windows.Media.TextOptions]::SetTextFormattingMode($root, [Windows.Media.TextFormattingMode]::Display)
        [Windows.Media.TextOptions]::SetTextRenderingMode($root, [Windows.Media.TextRenderingMode]::ClearType)
        [Windows.Media.TextOptions]::SetTextHintingMode($root, [Windows.Media.TextHintingMode]::Fixed)

        $layoutGrid = New-Object System.Windows.Controls.Grid
        $layoutGrid.Margin = New-Object Windows.Thickness(80, 60, 80, 60)
        $root.Children.Add($layoutGrid) | Out-Null

        $face = New-Object System.Windows.Controls.TextBlock
        $face.Text = ':('
        $face.FontSize = 190
        $face.FontFamily = 'Segoe UI'
        $face.Foreground = [Windows.Media.Brushes]::White
        $face.Margin = New-Object Windows.Thickness(0, 0, 0, 20)
        $face.HorizontalAlignment = 'Left'
        $face.VerticalAlignment = 'Top'
        $layoutGrid.Children.Add($face) | Out-Null

        $msg = New-Object System.Windows.Controls.TextBlock
        $msg.Text = "Your PC ran into a problem and needs to restart.`r`nWe're just collecting some error info, and then we'll restart for you."
        $msg.FontSize = 34
        $msg.FontFamily = 'Segoe UI'
        $msg.Foreground = [Windows.Media.Brushes]::White
        $msg.TextWrapping = 'Wrap'
        $msg.Margin = New-Object Windows.Thickness(0, 240, 0, 0)
        $msg.HorizontalAlignment = 'Left'
        $msg.VerticalAlignment = 'Top'
        $layoutGrid.Children.Add($msg) | Out-Null

        $percent = New-Object System.Windows.Controls.TextBlock
        $percent.Text = '0% complete'
        $percent.FontSize = 28
        $percent.FontFamily = 'Segoe UI'
        $percent.Foreground = [Windows.Media.Brushes]::White
        $percent.Margin = New-Object Windows.Thickness(0, 350, 0, 0)
        $percent.HorizontalAlignment = 'Left'
        $percent.VerticalAlignment = 'Top'
        $layoutGrid.Children.Add($percent) | Out-Null

        $stopcode = New-Object System.Windows.Controls.TextBlock
        $stopcode.Text = "For more information about this issue and possible fixes, visit https://www.windows.com/stopcode`r`n`r`nIf you call a support person, give them this info:`r`nStop code: CRITICAL_PROCESS_DIED"
        $stopcode.FontSize = 20
        $stopcode.FontFamily = 'Segoe UI'
        $stopcode.Foreground = [Windows.Media.Brushes]::White
        $stopcode.TextWrapping = 'Wrap'
        $stopcode.Margin = New-Object Windows.Thickness(0, 0, 0, 60)
        $stopcode.HorizontalAlignment = 'Left'
        $stopcode.VerticalAlignment = 'Bottom'
        $layoutGrid.Children.Add($stopcode) | Out-Null

        $viewbox.Child = $root

        $measureSize = New-Object System.Windows.Size($layoutWidth, $layoutHeight)
        $viewbox.Measure($measureSize)
        $viewbox.Arrange((New-Object System.Windows.Rect(0, 0, $layoutWidth, $layoutHeight)))
        $viewbox.UpdateLayout()

        $dpiX = 96.0 * $dpiScaleX
        $dpiY = 96.0 * $dpiScaleY
        $renderTarget = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($width, $height, $dpiX, $dpiY, [System.Windows.Media.PixelFormats]::Pbgra32)
        $renderTarget.Render($viewbox)

        $stride = $width * 4
        $pixelData = New-Object byte[] ($stride * $height)
        $renderTarget.CopyPixels($pixelData, $stride, 0)

        $bitmap = New-Object System.Drawing.Bitmap($width, $height, [System.Drawing.Imaging.PixelFormat]::Format32bppPArgb)
        $bitmapData = $bitmap.LockBits([System.Drawing.Rectangle]::new(0, 0, $width, $height), [System.Drawing.Imaging.ImageLockMode]::WriteOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppPArgb)
        [System.Runtime.InteropServices.Marshal]::Copy($pixelData, 0, $bitmapData.Scan0, $pixelData.Length)
        $bitmap.UnlockBits($bitmapData)

        return $bitmap
    }

    function New-BsodBitmap {
        $width = [int][Math]::Max(1, $screenBounds.Width)
        $height = [int][Math]::Max(1, $screenBounds.Height)

        $bitmap = New-Object System.Drawing.Bitmap($width, $height, [System.Drawing.Imaging.PixelFormat]::Format32bppPArgb)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([System.Drawing.Color]::FromArgb(0, 120, 215))

        foreach ($monitor in $script:monitorInfos) {
            $targetWidth = [int][Math]::Max(1, [int][Math]::Round($monitor.ActualWidth))
            $targetHeight = [int][Math]::Max(1, [int][Math]::Round($monitor.ActualHeight))
            $monitorBitmap = New-BsodSectionBitmap -targetWidth $targetWidth -targetHeight $targetHeight -scaleX $monitor.ScaleX -scaleY $monitor.ScaleY
            if ($monitorBitmap) {
                $offsetX = [int][Math]::Round($monitor.ActualLeft - $screenBounds.Left)
                $offsetY = [int][Math]::Round($monitor.ActualTop - $screenBounds.Top)
                $graphics.DrawImage($monitorBitmap, $offsetX, $offsetY, $targetWidth, $targetHeight)
                $monitorBitmap.Dispose()
            }
        }

        $graphics.Dispose()
        return $bitmap
    }

    function Start-ExitAnimation {
        if ($script:exitAnimationStarted) { return }

        if ($script:exitAnimationType -lt 1 -or $script:exitAnimationType -gt 4) {
            $script:exitAnimationType = 1
        }

        $script:pendingExitAnimation = $false
        $script:exitAnimationStarted = $true
        $script:exitAnimationStep = 0
        $script:exitAnimationTimer = [DateTime]::Now

        Clear-WindowCopies

        if ($script:exitAnimationType -eq 4) {
            if (-not $script:cursorHidden) {
                [System.Windows.Forms.Cursor]::Hide()
                $script:cursorHidden = $true
            }
            Clear-BsodGlitches -PreserveBeep:$script:bsodBeepScheduled
            Reset-Overlay
            $script:bsodRevealIndex = 0
            $script:bsodRevealTotal = 0
            $script:bsodRevealCompleted = $false
            $script:bsodSecondStageStartIndex = 0
            $bitmapWidth = [int][Math]::Max(1, $screenBounds.Width)
            $bitmapHeight = [int][Math]::Max(1, $screenBounds.Height)
            $script:overlayBitmap = New-Object System.Drawing.Bitmap($bitmapWidth, $bitmapHeight, [System.Drawing.Imaging.PixelFormat]::Format32bppPArgb)
            $pictureBox.Image = $script:overlayBitmap

            $script:bsodBitmap = New-BsodBitmap

            if (-not $script:bsodBeepProcess -or $script:bsodBeepProcess.HasExited) {
                Start-BsodBeep -DurationSeconds $BSOD_BEEP_DURATION_SECONDS
                $script:bsodBeepScheduled = $true
            }

            $graphics = [System.Drawing.Graphics]::FromImage($script:overlayBitmap)
            $graphics.DrawImage($script:bsodBitmap, 0, 0)
            $graphics.Dispose()
            $pictureBox.Invalidate()
        }
    }

$timer.Add_Tick({
    # Sprawdzenie ESC (kod 27) - natychmiastowe zakończenie
    if ([KeyboardHook]::GetAsyncKeyState(27) -ne 0) {
        if ($script:cursorHidden) {
            [System.Windows.Forms.Cursor]::Show()
            $script:cursorHidden = $false
        }
        $timer.Stop()
        $form.Close()
        return
    }

    if ($script:pendingExitAnimation -and -not $script:exitAnimationStarted) {
        if ($script:exitAnimationInitializing) {
            return
        }

        if ($script:coverCompleteAt -eq [DateTime]::MinValue) {
            $script:coverCompleteAt = [DateTime]::Now
        }

    $requiredDelay = if ($script:exitAnimationType -eq 4) { [Math]::Max(0.0, $BSOD_BEEP_LEAD_SECONDS) } else { 2.0 }
        if (([DateTime]::Now - $script:coverCompleteAt).TotalSeconds -lt $requiredDelay) {
            return
        }

        $script:exitAnimationInitializing = $true
        try {
            Start-ExitAnimation
        } finally {
            $script:exitAnimationInitializing = $false
        }
    }
    
    # Handle exit animation
    if ($script:exitAnimationStarted) {
        $elapsed = ([DateTime]::Now - $script:exitAnimationTimer).TotalSeconds
        
        switch ($script:exitAnimationType) {
            1 {
                Reset-Overlay
                Clear-WindowCopies
                Clear-BsodGlitches
                Set-NextGlitchSchedule
                return
            }
            2 {
                if ($script:exitAnimationStep -eq 0) {
                    Clear-OverlayRegion ([System.Drawing.Rectangle]::new(0, 0, $screenBounds.Width, [int]($screenBounds.Height / 2)))
                    $script:exitAnimationStep = 1
                    $script:exitAnimationTimer = [DateTime]::Now
                } elseif ($elapsed -ge 2.0) {
                    Reset-Overlay
                    Clear-WindowCopies
                    Clear-BsodGlitches
                    Set-NextGlitchSchedule
                    return
                }
            }
            3 {
                if ($script:exitAnimationStep -eq 0) {
                    Clear-OverlayRegion ([System.Drawing.Rectangle]::new(0, [int]($screenBounds.Height / 2), $screenBounds.Width, [int]([Math]::Ceiling($screenBounds.Height / 2))))
                    $script:exitAnimationStep = 1
                    $script:exitAnimationTimer = [DateTime]::Now
                } elseif ($elapsed -ge 2.0) {
                    Reset-Overlay
                    Clear-WindowCopies
                    Clear-BsodGlitches
                    Set-NextGlitchSchedule
                    return
                }
            }
            4 {
                if ($script:bsodPendingReveal) {
                    if ([DateTime]::Now -lt $script:bsodRevealDelayUntil) {
                        return
                    }

                    if ($script:overlayBitmap -and $script:bsodBitmap) {
                        $graphics = [System.Drawing.Graphics]::FromImage($script:overlayBitmap)
                        $graphics.DrawImage($script:bsodBitmap, 0, 0)
                        $graphics.Dispose()
                        $pictureBox.Invalidate()
                    }

                    $script:bsodPendingReveal = $false
                    $script:bsodRevealDelayUntil = [DateTime]::MinValue
                    $script:exitAnimationTimer = [DateTime]::Now
                    return
                }

                if ($script:exitAnimationStep -eq 0) {
                    if ($elapsed -lt 10.0) {
                        return
                    }

                    $script:exitAnimationStep = 1
                    $script:exitAnimationTimer = [DateTime]::Now

                    $glitchedBsod = Invoke-GlitchFilter -sourceBitmap $script:bsodBitmap -filterType 1
                    $script:bsodGlitchBitmaps.Add($glitchedBsod)

                    $script:windowCopies.Clear()
                    $vector = Get-RandomGlitchVector
                    $dirX = $vector.X
                    $dirY = $vector.Y

                    $stepDistance = [Math]::Max([Math]::Abs($dirX), [Math]::Abs($dirY))
                    $maxDimension = [double][Math]::Max($screenBounds.Width, $screenBounds.Height)
                    $numCopies = [Math]::Ceiling($maxDimension / $stepDistance) + 1

                    for ($i = 0; $i -lt $numCopies; $i++) {
                        $script:windowCopies.Add(@{
                            Bitmap = $glitchedBsod
                            OwnsBitmap = $false
                            X = [double]($screenBounds.Left + $dirX * $i)
                            Y = [double]($screenBounds.Top + $dirY * $i)
                            Width = $screenBounds.Width
                            Height = $screenBounds.Height
                            Visible = $false
                            Opacity = 0.75
                        })
                    }

                    $script:bsodRevealIndex = 0
                    $script:bsodRevealTotal = $script:windowCopies.Count
                    $script:bsodRevealCompleted = $false
                    $script:bsodSecondStageStartIndex = 0

                    Update-OverlayComposite -overlay $script:overlayBitmap -copies $script:windowCopies -baseBitmap $script:bsodBitmap
                    $pictureBox.Invalidate()
                    return
                }

                if ($script:exitAnimationStep -eq 1) {
                    if ($script:bsodRevealIndex -lt $script:bsodRevealTotal) {
                        $currentCopy = $script:windowCopies[$script:bsodRevealIndex]
                        $currentCopy.Visible = $true
                        Add-GlitchCopy -overlay $script:overlayBitmap -copy $currentCopy
                        $pictureBox.Invalidate()
                        $script:bsodRevealIndex++
                        return
                    }

                    if (-not $script:bsodRevealCompleted) {
                        $script:bsodRevealCompleted = $true
                        $script:exitAnimationTimer = [DateTime]::Now
                        return
                    }

                    if ($elapsed -lt 0.5) {
                        return
                    }

                    $script:exitAnimationStep = 2
                    $script:exitAnimationTimer = [DateTime]::Now

                    $secondGlitch = Invoke-GlitchFilter -sourceBitmap $script:bsodBitmap -filterType 1
                    $script:bsodGlitchBitmaps.Add($secondGlitch)

                    $vector = Get-RandomGlitchVector
                    $dirX = $vector.X
                    $dirY = $vector.Y

                    $stepDistance = [Math]::Max([Math]::Abs($dirX), [Math]::Abs($dirY))
                    $maxDimension = [double][Math]::Max($screenBounds.Width, $screenBounds.Height)
                    $numCopies = [Math]::Ceiling($maxDimension / $stepDistance) + 1
                    $startIndex = $script:windowCopies.Count

                    for ($i = 0; $i -lt $numCopies; $i++) {
                        $script:windowCopies.Add(@{
                            Bitmap = $secondGlitch
                            OwnsBitmap = $false
                            X = [double]($screenBounds.Left + $dirX * $i)
                            Y = [double]($screenBounds.Top + $dirY * $i)
                            Width = $screenBounds.Width
                            Height = $screenBounds.Height
                            Visible = $false
                            Opacity = 0.75
                        })
                    }

                    $script:bsodSecondStageStartIndex = $startIndex
                    $script:bsodRevealIndex = $script:bsodSecondStageStartIndex
                    $script:bsodRevealTotal = $script:windowCopies.Count
                    $script:bsodRevealCompleted = $false

                    return
                }

                if ($script:exitAnimationStep -eq 2) {
                    if ($script:bsodRevealIndex -lt $script:bsodRevealTotal) {
                        $currentCopy = $script:windowCopies[$script:bsodRevealIndex]
                        $currentCopy.Visible = $true
                        Add-GlitchCopy -overlay $script:overlayBitmap -copy $currentCopy
                        $pictureBox.Invalidate()
                        $script:bsodRevealIndex++
                        return
                    }

                    if (-not $script:bsodRevealCompleted) {
                        $script:bsodRevealCompleted = $true
                        $script:exitAnimationTimer = [DateTime]::Now
                        return
                    }

                    if ($elapsed -lt 1.5) {
                        return
                    }

                    Reset-Overlay
                    Clear-WindowCopies
                    Clear-BsodGlitches
                    Set-NextGlitchSchedule
                    return
                }
            }
        }
        return
    }

    # Check if waiting for next glitch
    if ($script:waitingForNextGlitch) {
        if ([DateTime]::Now -ge $script:nextGlitchTime) {
            $script:waitingForNextGlitch = $false
        } else {
            return
        }
    }
    
    # Start new glitch effect if no copies active
    if ($script:windowCopies.Count -eq 0 -and -not $script:exitAnimationStarted -and -not $script:waitingForNextGlitch -and -not $script:pendingExitAnimation) {
        $script:lastWindowCapture = [DateTime]::Now
        
        # Wyczyść stare kopie
        Clear-WindowCopies
        
        # Capture whole screen
        $screenBitmap = New-Object System.Drawing.Bitmap($screenBounds.Width, $screenBounds.Height)
        $screenGraphics = [System.Drawing.Graphics]::FromImage($screenBitmap)
        $screenGraphics.CopyFromScreen($screenBounds.Left, $screenBounds.Top, 0, 0, $screenBitmap.Size)
        $screenGraphics.Dispose()
        
        # Choose random filter type (1, 2, or 3)
        $filterType = $random.Next(1, 4)
        
        # Apply filter to create glitch effect bitmap
        $filteredBitmap = Invoke-GlitchFilter -sourceBitmap $screenBitmap -filterType $filterType
        $screenBitmap.Dispose()

        # Prepare overlay bitmap
        if ($script:overlayBitmap) {
            $pictureBox.Image = $null
            $script:overlayBitmap.Dispose()
        }
        $script:overlayBitmap = New-Object System.Drawing.Bitmap($screenBounds.Width, $screenBounds.Height, [System.Drawing.Imaging.PixelFormat]::Format32bppPArgb)
        $pictureBox.Image = $script:overlayBitmap
        
        # Wybierz losowy kierunek ruchu
        $vector = Get-RandomGlitchVector
        $dirX = $vector.X
        $dirY = $vector.Y
        
        # Oblicz liczbę kopii potrzebną, aby pokryć cały ekran (wzdłuż dłuższego boku)
        $stepDistance = [Math]::Max([Math]::Abs($dirX), [Math]::Abs($dirY))
        if ($stepDistance -le 0) {
            $stepDistance = 10
        }

        $maxDimension = [double][Math]::Max($screenBounds.Width, $screenBounds.Height)
        $numCopies = [Math]::Ceiling($maxDimension / $stepDistance) + 1
        if ($numCopies -lt 1) {
            $numCopies = 1
        }
        
        # Utwórz kopie całego ekranu z 75% opacity dla wszystkich
        for ($i = 0; $i -lt $numCopies; $i++) {
            # All copies use 75% opacity
            $opacity = 0.75
            $offsetX = $dirX * $i
            $offsetY = $dirY * $i
            
            $script:windowCopies.Add(@{
                Bitmap = $filteredBitmap
                OwnsBitmap = ($i -eq 0)
                X = [double]($screenBounds.Left + $offsetX)
                Y = [double]($screenBounds.Top + $offsetY)
                Width = $screenBounds.Width
                Height = $screenBounds.Height
                Visible = $false
                Opacity = $opacity
            })
        }
    }
    
    if ($script:windowCopies.Count -gt 0 -and -not $script:exitAnimationStarted) {
        if ($script:currentCopyIndex -lt $script:windowCopies.Count) {
            $currentCopy = $script:windowCopies[$script:currentCopyIndex]
            $currentCopy.Visible = $true
            if ($script:overlayBitmap) {
                Add-GlitchCopy -overlay $script:overlayBitmap -copy $currentCopy
                $pictureBox.Invalidate()
            }
            $script:currentCopyIndex++

            if (-not $script:pendingExitAnimation -and $script:currentCopyIndex -ge $script:windowCopies.Count) {
                $script:pendingExitAnimation = $true
                $script:coverCompleteAt = [DateTime]::Now
                $script:exitAnimationType = $random.Next(1, 5)
                if ($script:exitAnimationType -eq 4 -and -not $script:bsodBeepScheduled) {
                    if (-not $script:bsodBeepProcess -or $script:bsodBeepProcess.HasExited) {
                        Start-BsodBeep -DurationSeconds $BSOD_BEEP_DURATION_SECONDS
                    }
                    $script:bsodBeepScheduled = $true
                }
                Clear-WindowCopies
            }
        }
    }
})

$form.Add_Shown({
    $timer.Start()
})

$form.Add_FormClosing({
    $timer.Stop()
    Clear-WindowCopies
    Clear-BsodGlitches
    Reset-Overlay
    Stop-BsodBeep
    if ($script:cursorHidden) {
        [System.Windows.Forms.Cursor]::Show()
        $script:cursorHidden = $false
    }
})

$form.ShowDialog() | Out-Null

# Sprzątanie
$timer.Dispose()

if ($pictureBox.Image) {
    $pictureBox.Image.Dispose()
}
$form.Dispose()
Stop-BsodBeep

# Pokaż konsolę z powrotem
if ($consolePtr -ne [IntPtr]::Zero) {
    [Window]::ShowWindow($consolePtr, 5) | Out-Null
}