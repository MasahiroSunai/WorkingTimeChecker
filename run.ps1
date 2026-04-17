# =====================================
# run.ps1 - Step Driven Final Version
# =====================================

param(
    [string]$From,
    [string]$To,
    [string[]]$Only
)

$ErrorActionPreference = "Stop"

Write-Host "===== WorkingTimeChecker 開始 ====="

# -------------------------------------
# パス定義
# -------------------------------------
$rootDir   = $PSScriptRoot
$configDir = Join-Path $rootDir "config"
$envFile   = Join-Path $configDir ".env"
$configYml = Join-Path $configDir "config.yaml"

# -------------------------------------
# 1. .env 読み込み
# -------------------------------------
if (-not (Test-Path $envFile)) {
    throw ".env ファイルが見つかりません: $envFile"
}

Get-Content $envFile -Encoding UTF8 | ForEach-Object {
    $line = $_.Trim()

    if ($line -eq "" -or $line.StartsWith("#")) {
        return
    }

    $pair = $line -split "=", 2
    if ($pair.Count -ne 2) {
        return
    }

    $name  = $pair[0].Trim()
    $value = $pair[1].Trim()

    Set-Item -Path "env:$name" -Value $value
}

Write-Host "[OK] .env 読み込み完了"

# -------------------------------------
# 2. config.yaml 確認
# -------------------------------------
if (-not (Test-Path $configYml)) {
    throw "config.yaml が見つかりません: $configYml"
}

Write-Host "[OK] config.yaml 確認完了"

# -------------------------------------
# 3. Python 実行ユーティリティ
# -------------------------------------
function Run-Python {
    param (
        [string]$ScriptName
    )

    $scriptPath = Join-Path $rootDir $ScriptName
    if (-not (Test-Path $scriptPath)) {
        throw "Python スクリプトが見つかりません: $scriptPath"
    }

    Write-Host ">> 実行: $ScriptName"
    python $scriptPath --config $configYml

    if ($LASTEXITCODE -ne 0) {
        throw "$ScriptName の実行に失敗しました"
    }

    Write-Host "[OK] 完了: $ScriptName"
}

# -------------------------------------
# 4. 処理ステップ定義
# -------------------------------------
$Steps = @(
    @{
        Name   = "ConfluenceDownload"
        Action = { Run-Python "ConfluenceDownload.py" }
    },
    @{
        Name   = "ExcelMacro"
        Action = {
            Write-Host ">> 実行: Run-ExcelMacro.ps1"
            & "$rootDir\Run-ExcelMacro.ps1"

            if ($LASTEXITCODE -ne 0) {
                throw "Run-ExcelMacro.ps1 の実行に失敗しました"
            }

            Write-Host "[OK] 完了: Run-ExcelMacro.ps1"
        }
    },
    @{
        Name   = "WebAttendanceDownload"
        Action = { Run-Python "WebAttendanceDownload.py" }
    },
    @{
        Name   = "WorkingTimeChecker"
        Action = { Run-Python "WorkingTimeChecker.py" }
    },
    @{
        Name   = "CopyConfluence"
        Action = { Run-Python "CopyConfluence.py" }
    },
    @{
        Name   = "PostRocketChat"
        Action = {
            Write-Host ">> 実行: PostRocketChatMessage.py"

            python "$rootDir\PostRocketChatMessage.py" `
                --config "$configYml" `
                --message "@all`n工数表⇔Web勤怠チェックが完了しました。Confluence を確認してください。"

            if ($LASTEXITCODE -ne 0) {
                throw "PostRocketChatMessage.py の実行に失敗しました"
            }

            Write-Host "[OK] 完了: PostRocketChatMessage.py"
        }
    }
)

# -------------------------------------
# 5. ステップ実行制御
# -------------------------------------
try {
    $run = $false

    foreach ($step in $Steps) {

        # Only 指定がある場合（指定ステップのみ実行）
        if ($Only) {
            if ($Only -contains $step.Name) {
                Write-Host ">> ステップ実行: $($step.Name)"
                & $step.Action
            }
            continue
        }

        # From 判定
        if ($From -and $step.Name -eq $From) {
            $run = $true
        }

        if (-not $From) {
            $run = $true
        }

        if ($run) {
            Write-Host ">> ステップ実行: $($step.Name)"
            & $step.Action
        }

        # To 判定
        if ($To -and $step.Name -eq $To) {
            break
        }
    }

    Write-Host "===== [OK] 全処理 正常完了 ====="
}
catch {
    Write-Error "[NG] エラー発生: $_"
    exit 1
}