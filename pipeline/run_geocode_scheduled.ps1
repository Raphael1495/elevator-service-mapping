# Windows 작업 스케줄러가 이 스크립트를 호출한다.
# 예약 실행이 원인 불명으로 조용히 실패한 적이 있어서(로그에 아무것도 안 남음),
# 무슨 일이 있어도 실행 시점/환경 정보만큼은 반드시 남기도록 만든 진단용 래퍼.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$debugLog = Join-Path $PSScriptRoot "scheduler_debug.log"
$runLog = Join-Path $PSScriptRoot "geocode_run.log"

function Write-DebugLine($msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $debugLog -Value $line -Encoding UTF8
}

Write-DebugLine "===== 예약 실행 시작 ====="
Write-DebugLine ("실행 사용자: " + [System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
Write-DebugLine ("작업 폴더: " + $root)
Write-DebugLine ("전원 상태: " + (Get-WmiObject Win32_Battery -ErrorAction SilentlyContinue | Select-Object -ExpandProperty BatteryStatus))

$addressFile = Join-Path $root "국가DB_SID_주소.xlsx"
Write-DebugLine ("주소 원본 파일 경로: " + $addressFile)
Write-DebugLine ("주소 원본 파일 존재 여부: " + (Test-Path -LiteralPath $addressFile))

$configFile = Join-Path $root "config.local.json"
Write-DebugLine ("config.local.json 존재 여부: " + (Test-Path -LiteralPath $configFile))

$pythonExe = "C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe"
Write-DebugLine ("python.exe 존재 여부: " + (Test-Path -LiteralPath $pythonExe))

try {
    & $pythonExe (Join-Path $PSScriptRoot "geocode.py") *>> $runLog
    Write-DebugLine ("geocode.py 종료 코드: " + $LASTEXITCODE)
} catch {
    Write-DebugLine ("예외 발생: " + $_.Exception.Message)
}

Write-DebugLine "===== 예약 실행 종료 ====="
