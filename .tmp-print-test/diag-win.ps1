$ErrorActionPreference = "Stop"
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class WinEnum {
  public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc lpEnumFunc, IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hWnd, EnumProc lpEnumFunc, IntPtr lParam);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int max);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder sb, int max);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
}
"@
function Janelas($pidAlvo) {
    $script:linhas = @()
    $callback = [WinEnum+EnumProc]{
        param($hWnd, $lParam)
        $procId = [uint32]0
        [WinEnum]::GetWindowThreadProcessId($hWnd, [ref]$procId) | Out-Null
        if ($procId -eq $pidAlvo) {
            $sb = New-Object System.Text.StringBuilder 512
            [WinEnum]::GetWindowText($hWnd, $sb, $sb.Capacity) | Out-Null
            $cls = New-Object System.Text.StringBuilder 256
            [WinEnum]::GetClassName($hWnd, $cls, $cls.Capacity) | Out-Null
            $vis = [WinEnum]::IsWindowVisible($hWnd)
            $script:linhas += ("vis={0} class={1} title={2}" -f $vis, $cls, $sb)
        }
        return $true
    }
    [WinEnum]::EnumWindows($callback, [IntPtr]::Zero) | Out-Null
    $script:linhas | ForEach-Object { Write-Output $_ }
}
$app = New-Object -ComObject Excel.Application
Start-Sleep -Seconds 2
Write-Output "READY1 $($app.Ready) VIS $($app.Visible)"
$pids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
Write-Output "PIDS $($pids -join ',')"
foreach ($p in $pids) { Write-Output "-- pid $p --"; Janelas $p }
$app.Visible = $false
$app.DisplayAlerts = $false
$doc = $app.Workbooks.Open("D:\AtestoFacil\.tmp-print-test\01-previsao.xlsx", 0, $true)
Start-Sleep -Seconds 2
Write-Output "READY2 $($app.Ready)"
foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })) { Write-Output "-- pid $p after open --"; Janelas $p }
try { $app.Quit() } catch { Write-Output $_.Exception.Message }
Start-Sleep -Seconds 1
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
