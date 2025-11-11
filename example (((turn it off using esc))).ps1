Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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
[Window]::ShowWindow($consolePtr, 0) | Out-Null

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

# Ustawienie formularza jako overlay
$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = 'None'
$form.StartPosition = 'Manual'
$form.Location = New-Object System.Drawing.Point($screenBounds.Left, $screenBounds.Top)
$form.Size = New-Object System.Drawing.Size($screenBounds.Width, $screenBounds.Height)
$form.TopMost = $true
$form.BackColor = [System.Drawing.Color]::Black
$form.TransparencyKey = [System.Drawing.Color]::Black
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
$pictureBox.BackColor = [System.Drawing.Color]::Black
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Normal
$form.Controls.Add($pictureBox)

$random = New-Object System.Random

# Lista animowanych fragmentów okien
$animatedFragments = New-Object System.Collections.Generic.List[object]

# Timer do animacji - 100 FPS (10ms)
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 10
$frameCounter = 0

$timer.Add_Tick({
    # Sprawdzenie ESC (kod 27)
    if ([KeyboardHook]::GetAsyncKeyState(27) -ne 0) {
        $timer.Stop()
        $form.Close()
        return
    }

    $script:frameCounter++
    
    # Co kilka klatek dodaj nowe animowane fragmenty
    if ($script:frameCounter % 3 -eq 0) {
        # Przechwycenie screenshota ekranu
        $screenshot = New-Object System.Drawing.Bitmap($screenBounds.Width, $screenBounds.Height)
        $graphics = [System.Drawing.Graphics]::FromImage($screenshot)
        $graphics.CopyFromScreen($screenBounds.Left, $screenBounds.Top, 0, 0, $screenshot.Size)
        $graphics.Dispose()
        
        # Wybierz losowy prostokąt z ekranu
        $srcX = $random.Next(0, [Math]::Max(1, $screenBounds.Width - 400))
        $srcY = $random.Next(0, [Math]::Max(1, $screenBounds.Height - 300))
        $width = $random.Next(250, 600)
        $height = $random.Next(200, 500)
        
        # Ogranicz do granic ekranu
        if ($srcX + $width -gt $screenBounds.Width) {
            $width = $screenBounds.Width - $srcX
        }
        if ($srcY + $height -gt $screenBounds.Height) {
            $height = $screenBounds.Height - $srcY
        }
        
        if ($width -gt 50 -and $height -gt 50) {
            # Skopiuj fragment do oddzielnej bitmapy
            $fragmentBitmap = New-Object System.Drawing.Bitmap($width, $height)
            $fragmentGraphics = [System.Drawing.Graphics]::FromImage($fragmentBitmap)
            $fragmentGraphics.DrawImage($screenshot, 0, 0, (New-Object System.Drawing.Rectangle($srcX, $srcY, $width, $height)), [System.Drawing.GraphicsUnit]::Pixel)
            $fragmentGraphics.Dispose()
            
            # Wybierz losowy kierunek ruchu
            $dirX = $random.Next(-3, 4)
            $dirY = $random.Next(-3, 4)
            
            # Upewnij się że jest jakiś ruch
            if ($dirX -eq 0 -and $dirY -eq 0) {
                $dirX = $random.Next(0, 2) * 2 - 1
            }
            
            # Utwórz 8-12 kopii tego samego fragmentu
            $numCopies = $random.Next(8, 13)
            
            for ($i = 0; $i -lt $numCopies; $i++) {
                # Każda kopia zaczyna w tym samym miejscu ale z małym offsetem
                $offsetMultiplier = $i * 2
                
                $script:animatedFragments.Add(@{
                    Bitmap = $fragmentBitmap
                    X = [double]($srcX + $dirX * $offsetMultiplier)
                    Y = [double]($srcY + $dirY * $offsetMultiplier)
                    DirectionX = $dirX
                    DirectionY = $dirY
                    Life = 180  # ~1.8s życia
                    Width = $width
                    Height = $height
                })
            }
        }
        
        $screenshot.Dispose()
    }
    
    # Tworzenie obrazu z nałożonymi fragmentami
    $glitchBitmap = New-Object System.Drawing.Bitmap($screenBounds.Width, $screenBounds.Height)
    $glitchGraphics = [System.Drawing.Graphics]::FromImage($glitchBitmap)
    $glitchGraphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighSpeed
    $glitchGraphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighSpeed
    $glitchGraphics.Clear([System.Drawing.Color]::Transparent)
    
    # Rysuj i animuj wszystkie fragmenty
    $toRemove = New-Object System.Collections.Generic.List[int]
    
    for ($i = 0; $i -lt $script:animatedFragments.Count; $i++) {
        $frag = $script:animatedFragments[$i]
        
        # Rysuj fragment z 20% przezroczystością
        $colorMatrix = New-Object System.Drawing.Imaging.ColorMatrix
        $colorMatrix.Matrix33 = 0.2  # 20% przezroczystości
        
        $imageAttributes = New-Object System.Drawing.Imaging.ImageAttributes
        $imageAttributes.SetColorMatrix($colorMatrix)
        
        $destRect = New-Object System.Drawing.Rectangle([int]$frag.X, [int]$frag.Y, $frag.Width, $frag.Height)
        
        try {
            if ($destRect.X -gt -$frag.Width -and $destRect.X -lt $screenBounds.Width -and
                $destRect.Y -gt -$frag.Height -and $destRect.Y -lt $screenBounds.Height) {
                $glitchGraphics.DrawImage(
                    $frag.Bitmap,
                    $destRect,
                    0, 0, $frag.Width, $frag.Height,
                    [System.Drawing.GraphicsUnit]::Pixel,
                    $imageAttributes
                )
            }
        } catch {}
        
        $imageAttributes.Dispose()
        
        # Przesuń fragment w tym samym kierunku
        $frag.X += $frag.DirectionX
        $frag.Y += $frag.DirectionY
        
        # Zmniejsz życie
        $frag.Life--
        
        if ($frag.Life -le 0) {
            $toRemove.Add($i)
        }
    }
    
    # Usuń "martwe" fragmenty
    for ($i = $toRemove.Count - 1; $i -ge 0; $i--) {
        $idx = $toRemove[$i]
        # Bitmap jest współdzielona przez wiele kopii, usuń tylko raz
        $bitmapToCheck = $script:animatedFragments[$idx].Bitmap
        $script:animatedFragments.RemoveAt($idx)
        
        # Sprawdź czy to była ostatnia kopia używająca tej bitmapy
        $stillUsed = $false
        foreach ($frag in $script:animatedFragments) {
            if ([object]::ReferenceEquals($frag.Bitmap, $bitmapToCheck)) {
                $stillUsed = $true
                break
            }
        }
        
        if (-not $stillUsed) {
            $bitmapToCheck.Dispose()
        }
    }
    
    # Ogranicz liczbę fragmentów
    while ($script:animatedFragments.Count -gt 200) {
        $oldestBitmap = $script:animatedFragments[0].Bitmap
        $script:animatedFragments.RemoveAt(0)
        
        # Sprawdź czy bitmap jest jeszcze używana
        $stillUsed = $false
        foreach ($frag in $script:animatedFragments) {
            if ([object]::ReferenceEquals($frag.Bitmap, $oldestBitmap)) {
                $stillUsed = $true
                break
            }
        }
        
        if (-not $stillUsed) {
            $oldestBitmap.Dispose()
        }
    }
    
    # Aktualizacja obrazu
    if ($pictureBox.Image) {
        $oldImage = $pictureBox.Image
        $pictureBox.Image = $null
        $oldImage.Dispose()
    }
    $pictureBox.Image = $glitchBitmap
    
    # Sprzątanie
    $glitchGraphics.Dispose()
})

$form.Add_Shown({
    $timer.Start()
})

$form.Add_FormClosing({
    $timer.Stop()
    
    # Sprzątanie fragmentów
    $uniqueBitmaps = New-Object System.Collections.Generic.HashSet[object]
    foreach ($frag in $animatedFragments) {
        if ($frag.Bitmap) {
            $uniqueBitmaps.Add($frag.Bitmap) | Out-Null
        }
    }
    foreach ($bmp in $uniqueBitmaps) {
        $bmp.Dispose()
    }
    $animatedFragments.Clear()
})

$form.ShowDialog() | Out-Null

# Sprzątanie
$timer.Dispose()

if ($pictureBox.Image) {
    $pictureBox.Image.Dispose()
}
$form.Dispose()

# Pokaż konsolę z powrotem
if ($consolePtr -ne [IntPtr]::Zero) {
    [Window]::ShowWindow($consolePtr, 5) | Out-Null
}
