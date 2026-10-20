$ErrorActionPreference = "Stop"
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class WinAct {
  public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc lpEnumFunc, IntPtr lParam);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int max);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder sb, int max);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
}
"@
function Fechar-Ativacao([int[]]$Pids) {
    $fechou = $false
    $callback = [WinAct+EnumProc]{
        param($hWnd, $lParam)
        $procId = [uint32]0
        [WinAct]::GetWindowThreadProcessId($hWnd, [ref]$procId) | Out-Null
        if ($Pids -contains [int]$procId) {
            $cls = New-Object System.Text.StringBuilder 256
            [WinAct]::GetClassName($hWnd, $cls, $cls.Capacity) | Out-Null
            $sb = New-Object System.Text.StringBuilder 512
            [WinAct]::GetWindowText($hWnd, $sb, $sb.Capacity) | Out-Null
            $titulo = $sb.ToString()
            $classe = $cls.ToString()
            if ($classe -eq "NUIDialog" -and ($titulo -match "Ativ|Activ")) {
                [WinAct]::PostMessage($hWnd, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
                $script:achou = $true
            }
        }
        return $true
    }
    $script:achou = $false
    [WinAct]::EnumWindows($callback, [IntPtr]::Zero) | Out-Null
    return $script:achou
}
function Exportar($origem, $destino, $rotulo) {
    Write-Output "---- $rotulo ----"
    $antes = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
    $app = New-Object -ComObject Excel.Application
    Start-Sleep -Milliseconds 800
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $app.AskToUpdateLinks = $false
    $doc = $app.Workbooks.Open($origem, 0, $true)
    $novos = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id } | Where-Object { $antes -notcontains $_ })
    $pronto = $false
    for ($i = 0; $i -lt 25; $i++) {
        $null = Fechar-Ativacao $novos
        try { if ($app.Ready) { $pronto = $true; break } } catch {}
        Start-Sleep -Milliseconds 400
    }
    Write-Output ("READY " + $app.Ready + " pronto=" + $pronto)
    try {
        $doc.ExportAsFixedFormat(0, $destino)
        Write-Output "OK export"
    } catch {
        Write-Output ("FAIL export :: " + $_.Exception.Message)
    }
    try { $doc.Close($false) } catch {}
    try { $app.Quit() } catch { Write-Output ("FAIL quit :: " + $_.Exception.Message) }
    $limite = (Get-Date).AddSeconds(10)
    do {
        $vivos = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $antes -notcontains [int]$_.Id })
        if ($vivos.Count -eq 0) { break }
        Start-Sleep -Milliseconds 400
    } while ((Get-Date) -lt $limite)
    foreach ($item in $vivos) { Stop-Process -Id $item.Id -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Milliseconds 800
    if (Test-Path $destino) { Write-Output ("SIZE " + (Get-Item $destino).Length) } else { Write-Output "NOFILE" }
}
Exportar "D:\AtestoFacil\.tmp-print-test\01-previsao.xlsx" "D:\AtestoFacil\.tmp-print-test\pdf\a-previsao.pdf" "previsao"
Exportar "D:\AtestoFacil\.tmp-print-test\03-rpa.xlsx" "D:\AtestoFacil\.tmp-print-test\pdf\c-rpa.pdf" "rpa"
