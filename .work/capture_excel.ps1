param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$OutputDir
)

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class WinCapture {
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$excel.DisplayFormulaBar = $true
$excel.DisplayStatusBar = $true
$excel.WindowState = -4137
$workbook = $excel.Workbooks.Open($WorkbookPath, 0, $true)
$shell = New-Object -ComObject WScript.Shell

function Save-ExcelWindow([string]$name) {
    Start-Sleep -Milliseconds 800
    $hwnd = [IntPtr]$excel.Hwnd
    [WinCapture]::ShowWindow($hwnd, 3) | Out-Null
    [WinCapture]::SetForegroundWindow($hwnd) | Out-Null
    Start-Sleep -Milliseconds 500
    $rect = New-Object WinCapture+RECT
    [WinCapture]::GetWindowRect($hwnd, [ref]$rect) | Out-Null
    $width = $rect.Right - $rect.Left
    $height = $rect.Bottom - $rect.Top
    $bitmap = New-Object System.Drawing.Bitmap($width, $height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, $bitmap.Size)
    $path = Join-Path $OutputDir ($name + '.png')
    $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $bitmap.Dispose()
}

$shots = @(
    @{ SheetIndex=1; Cell='F7'; Name='01_vlookup_exact'; Edit=$true },
    @{ SheetIndex=2; Cell='E7'; Name='02_vlookup_approx'; Edit=$true },
    @{ SheetIndex=3; Cell='I7'; Name='03_hlookup'; Edit=$true },
    @{ SheetIndex=4; Cell='E7'; Name='04_choose'; Edit=$true },
    @{ SheetIndex=5; Cell='F8'; Name='05_index_match'; Edit=$true },
    @{ SheetIndex=6; Cell='J5'; Name='06_practice'; Edit=$false },
    @{ SheetIndex=7; Cell='B5'; Name='07_answers'; Edit=$false }
)

foreach ($shot in $shots) {
    $ws = $workbook.Worksheets.Item($shot.SheetIndex)
    $ws.Activate()
    $excel.ActiveWindow.DisplayGridlines = $true
    $excel.ActiveWindow.DisplayHeadings = $true
    $excel.ActiveWindow.Zoom = 100
    $excel.ActiveWindow.ScrollRow = 1
    $excel.ActiveWindow.ScrollColumn = 1
    $excel.Goto($ws.Range($shot.Cell), $false)
    [WinCapture]::SetForegroundWindow([IntPtr]$excel.Hwnd) | Out-Null
    $shell.AppActivate($excel.Caption) | Out-Null
    Start-Sleep -Milliseconds 400
    if ($shot.Edit) {
        $shell.SendKeys('{F2}')
        Start-Sleep -Milliseconds 500
    }
    Save-ExcelWindow $shot.Name
    if ($shot.Edit) {
        $shell.SendKeys('{ESC}')
        Start-Sleep -Milliseconds 250
    }
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
