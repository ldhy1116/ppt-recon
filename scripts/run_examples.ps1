# run_examples.ps1 - PPT 重组一键示例脚本（测试用例 test1/test2/test3）
# 用法：
#   1. 确保 Ollama 运行中（ollama serve）；未运行则自动回退关键词模式（0 token）
#   2. PowerShell 中执行：.\scripts\run_examples.ps1
# 说明：依次对三个用例执行重组，产物输出到 data\。
#       Ollama 连接默认值已内置在 scripts\pptx_reorganize_cli.py，无需设置环境变量。

$ErrorActionPreference = "Continue"

# 工作目录设为项目根目录（脚本位于 scripts\ 下）
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir = Split-Path -Parent $scriptDir
Set-Location $rootDir
Write-Host "=== 工作目录：$rootDir ===" -ForegroundColor Cyan
Write-Host "=== 大模型：Ollama qwen2.5:3b（CLI 内置默认值，未启动时自动回退关键词模式）===" -ForegroundColor Cyan
Write-Host ""

# 1. 测试任务清单（test1 研究报告18页 / test2 课件57页第N节 / test3 招生宣讲60页两级编号）
$tasks = @(
    @{ name = "test1 研究报告 → 技术汇报";       pptx = "ppts\test1.pptx"; purpose = "技术汇报";     output = "data\test1_tech.pptx" },
    @{ name = "test2 课件第N节 → 课堂教学";       pptx = "ppts\test2.pptx"; purpose = "课堂教学";     output = "data\test2_class.pptx" },
    @{ name = "test3 两级编号 → 招生综合宣讲";    pptx = "ppts\test3.pptx"; purpose = "招生综合宣讲"; output = "data\test3_intro.pptx" }
)

# 2. 依次执行
$success = 0
$failed  = 0
foreach ($task in $tasks) {
    Write-Host "=== $($task.name) ===" -ForegroundColor Yellow
    $cmd = "python scripts\pptx_reorganize_cli.py reorganize `"$($task.pptx)`" --purpose `"$($task.purpose)`" -o `"$($task.output)`" --smart --use-llm --yes --force"
    Write-Host "CMD: $cmd" -ForegroundColor Gray
    try {
        Invoke-Expression $cmd
        if ($LASTEXITCODE -eq 0) {
            Write-Host "OK: $($task.output)" -ForegroundColor Green
            $success++
        } else {
            Write-Host "FAIL (exit $LASTEXITCODE)" -ForegroundColor Red
            $failed++
        }
    } catch {
        Write-Host "ERROR: $_" -ForegroundColor Red
        $failed++
    }
    Write-Host ""
}

# 汇总
Write-Host "=== 执行完成 ===" -ForegroundColor Cyan
Write-Host "成功：$success / $($tasks.Count)"
if ($failed -gt 0) { Write-Host "失败：$failed" -ForegroundColor Red }
Write-Host "产物目录：$rootDir\data\" -ForegroundColor Cyan
