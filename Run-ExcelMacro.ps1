# =====================================
# Run-ExcelMacro.ps1
# Excel VBA マクロ実行（PowerShell対応版）
# =====================================

$ErrorActionPreference = "Stop"

function Retry-ExcelAction {
    param (
        [scriptblock]$Action,
        [int]$Retry = 10,
        [int]$SleepMs = 500
    )

    for ($i = 1; $i -le $Retry; $i++) {
        try {
            & $Action
            return
        } catch {
            Write-Host "Retry $i / $Retry : $($_.Exception.Message)"
            if ($i -eq $Retry) { throw }
            Start-Sleep -Milliseconds ($SleepMs * $i)
        }
    }
}

Write-Host ">> Excel マクロ実行開始"

# -------------------------------------
# 1. パス解決
# -------------------------------------
$rootDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$configYml = Join-Path $rootDir "config\config.yaml"

if (-not (Test-Path $configYml)) {
    Write-Error "config.yaml が見つかりません: $configYml"
    exit 1
}


# -------------------------------------
# 2. config.yaml 読み込み（python -c / PYTHONPATH対応）
# -------------------------------------
try {
    $env:PYTHONPATH = $rootDir

    $json = python -c "
from utils.config_loader import load_config
from utils.month_utils import resolve_target_month
from utils.path_utils import expand_path
import json

config = load_config(r'$configYml')
vars = resolve_target_month(config)

excelPath = expand_path(
    config['paths']['workload_aggregate_file'],
    vars
)

macros = config.get('excel', {}).get('macros', [])

print(json.dumps({
    'excelPath': excelPath,
    'macros': macros
}))
"
}
catch {
    Write-Error "config.yaml の読み込みに失敗しました: $_"
    exit 1
}
finally {
    Remove-Item Env:PYTHONPATH -ErrorAction SilentlyContinue
}


$data = $json | ConvertFrom-Json

$excelPath = $data.excelPath
$macroCandidates = @()

foreach ($m in $data.macros) {
    $macroCandidates += [string]$m
}

if ($macroCandidates.Count -eq 0) {
    throw "config.yaml に excel.macros が定義されていません"
}

if (-not (Test-Path $excelPath)) {
    Write-Error "Excel ファイルが存在しません: $excelPath"
    exit 1
}

Write-Host "[OK] 対象 Excel: $excelPath"

# -------------------------------------
# 3. Excel COM 起動
# -------------------------------------
$excel = $null
$workbook = $null
$excelWasCreated = $false

try {
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        Write-Host "[OK] 既存の Excel を使用"
    } catch {
        $excel = New-Object -ComObject Excel.Application
        $excelWasCreated = $true
        Write-Host "[OK] 新しい Excel を起動"
    }

    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.ScreenUpdating = $false } catch {}

    # ---------------------------------
    # 4. Workbook を開く
    # ---------------------------------
    $workbook = $excel.Workbooks.Open($excelPath)
    if (-not $workbook) {
        throw "Workbook を開けませんでした"
    }

    Start-Sleep -Milliseconds 500

    # ---------------------------------
    # 5. VBA マクロ実行（config.yaml 参照）
    # ---------------------------------
    Write-Host ">> マクロ実行開始（数分かかる場合があります）"
    $startTime = Get-Date

    $ran = $false
    foreach ($macro in $macroCandidates) {
        if ($ran) { break }

        Write-Host ">> マクロ実行試行: $macro"
        for ($i = 1; $i -le 10; $i++) {
            try {
                $excel.Run($macro)
                Write-Host "[OK] マクロ成功: $macro"
                $ran = $true
                break
            } catch {
                Start-Sleep -Milliseconds (300 + 200 * $i)
            }
        }
    }

    if (-not $ran) {
        throw "VBA マクロを実行できませんでした（config.yaml の macro 名を確認してください）"
    }
    else {
        $elapsed = New-TimeSpan -Start $startTime -End (Get-Date)
        Write-Host "[OK] マクロ実行完了: $($elapsed.TotalMinutes) 分"
    }
    Start-Sleep -Seconds 3

    # ---------------------------------
    # 6. 保存
    # ---------------------------------
    try {
        $xlOpenXMLWorkbookMacroEnabled = 52
        Retry-ExcelAction { $workbook.SaveAs($excelPath, $xlOpenXMLWorkbookMacroEnabled) }
    } catch {
        Retry-ExcelAction { $workbook.Save() }
    }

    Write-Host "[OK] Excel 保存完了"
}
catch {
    Write-Error "[NG] Excel マクロ実行中にエラー: $_"
    exit 1
}
finally {
    if ($workbook) {
        Retry-ExcelAction { $workbook.Close($false) }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }

    if ($excelWasCreated -and $excel) {
        Retry-ExcelAction { $excel.Quit() }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "[OK] Excel マクロ実行完了"
exit 0