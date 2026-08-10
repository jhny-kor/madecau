[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path (Get-Location) "captures\markany-ui.json"),
    [ValidateRange(1, 30)]
    [int]$DelaySeconds = 5,
    [ValidateRange(1, 15)]
    [int]$MaxDepth = 8,
    [switch]$SelfTest
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

if ($SelfTest) {
    if (-not ([System.Windows.Automation.AutomationElement] -as [type])) {
        throw "Windows UI Automation is unavailable."
    }
    Write-Output "self-test: ok"
    exit 0
}

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class MarkAnyNativeWindow
{
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
}
"@

Write-Output "Focus the MarkAny window. Capturing in $DelaySeconds seconds..."
Start-Sleep -Seconds $DelaySeconds

$handle = [MarkAnyNativeWindow]::GetForegroundWindow()
if ($handle -eq [IntPtr]::Zero) {
    throw "No foreground window was found."
}

$root = [System.Windows.Automation.AutomationElement]::FromHandle($handle)
if ($null -eq $root) {
    throw "The foreground window does not expose a UI Automation tree."
}

$queue = [System.Collections.Generic.Queue[object]]::new()
$queue.Enqueue([pscustomobject]@{ Element = $root; Depth = 0 })
$controls = [System.Collections.Generic.List[object]]::new()

while ($queue.Count -gt 0) {
    $item = $queue.Dequeue()
    $element = $item.Element
    $depth = $item.Depth

    try {
        $current = $element.Current
        $rectangle = $current.BoundingRectangle
        $controls.Add([pscustomobject]@{
            depth = $depth
            controlType = $current.ControlType.ProgrammaticName.Replace("ControlType.", "")
            name = $current.Name
            automationId = $current.AutomationId
            className = $current.ClassName
            frameworkId = $current.FrameworkId
            enabled = $current.IsEnabled
            offscreen = $current.IsOffscreen
            nativeWindowHandle = $current.NativeWindowHandle
            bounds = [pscustomobject]@{
                x = $rectangle.X
                y = $rectangle.Y
                width = $rectangle.Width
                height = $rectangle.Height
            }
        })

        if ($depth -ge $MaxDepth) {
            continue
        }

        $children = $element.FindAll(
            [System.Windows.Automation.TreeScope]::Children,
            [System.Windows.Automation.Condition]::TrueCondition
        )
        foreach ($child in $children) {
            $queue.Enqueue([pscustomobject]@{ Element = $child; Depth = $depth + 1 })
        }
    }
    catch [System.Windows.Automation.ElementNotAvailableException] {
        continue
    }
}

$processId = $root.Current.ProcessId
$process = Get-Process -Id $processId
$capture = [pscustomobject]@{
    capturedAt = [DateTimeOffset]::Now.ToString("o")
    processId = $processId
    processName = $process.ProcessName
    windowTitle = $root.Current.Name
    maxDepth = $MaxDepth
    controls = $controls
}

$absoluteOutputPath = [IO.Path]::GetFullPath($OutputPath)
$outputDirectory = [IO.Path]::GetDirectoryName($absoluteOutputPath)
if (-not [string]::IsNullOrEmpty($outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}
$capture | ConvertTo-Json -Depth 5 | Set-Content -Path $absoluteOutputPath -Encoding UTF8

Write-Output "captured: $absoluteOutputPath"
Write-Output "controls: $($controls.Count)"
Write-Output "Review names and the window title for sensitive information before sharing."
