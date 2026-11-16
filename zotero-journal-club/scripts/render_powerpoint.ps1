param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$taskInput = [IO.Path]::GetFullPath($InputPath)
$taskOutput = [IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path -LiteralPath $taskInput -PathType Leaf) -or [IO.Path]::GetExtension($taskInput) -ine '.pptx') {
    throw 'Input must be an existing PPTX file.'
}
if ([IO.Path]::GetExtension($taskOutput) -ine '.pdf' -or $taskInput -ieq $taskOutput) {
    throw 'Output must be a separate PDF file.'
}
if (Test-Path -LiteralPath $taskOutput) {
    throw "Refusing to overwrite existing output: $taskOutput"
}

$taskPowerPoint = $null
$taskPresentation = $null
$taskCreatedApplication = $false
$taskSecurity = $null
try {
    $taskRunning = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
    if ($taskRunning) {
        # An inaccessible active app must not be replaced or quit by this helper.
        $taskPowerPoint = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    }
    else {
        $taskPowerPoint = New-Object -ComObject PowerPoint.Application
        $taskCreatedApplication = $true
    }
    $taskSecurity = $taskPowerPoint.AutomationSecurity
    $taskPowerPoint.AutomationSecurity = 3
    # WithWindow=false opens only this presentation without hiding the application.
    $taskPresentation = $taskPowerPoint.Presentations.Open($taskInput, -1, 0, 0)
    $taskPowerPoint.AutomationSecurity = $taskSecurity
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($taskOutput)) | Out-Null
    $taskPresentation.ExportAsFixedFormat($taskOutput, 2)
    if (-not (Test-Path -LiteralPath $taskOutput -PathType Leaf) -or (Get-Item -LiteralPath $taskOutput).Length -eq 0) {
        throw 'PowerPoint did not create a nonempty PDF.'
    }
    Write-Output "Exported PDF: $taskOutput"
}
finally {
    if ($null -ne $taskPresentation) {
        try { $taskPresentation.Close() } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskPresentation)
    }
    if ($null -ne $taskPowerPoint) {
        if ($null -ne $taskSecurity) {
            try { $taskPowerPoint.AutomationSecurity = $taskSecurity } catch {}
        }
        if ($taskCreatedApplication) {
            try {
                if ($taskPowerPoint.Presentations.Count -eq 0) { $taskPowerPoint.Quit() }
            } catch {}
        }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskPowerPoint)
    }
}
