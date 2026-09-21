@echo off
setlocal EnableExtensions
title Publisher to PDF Converter v7
set "LISTFILE=%~1"

echo.
echo ============================================
echo       Publisher to PDF Converter
echo ============================================
echo.
if defined LISTFILE goto introlist
echo This will scan this folder and ALL subfolders:
echo.
echo "%~dp0"
goto introdone
:introlist
echo Only the files named in this list will be converted:
echo.
echo "%LISTFILE%"
:introdone
echo.
echo IMPORTANT: Please close Microsoft Publisher first.
echo.
echo If it seems stuck, look on the taskbar for a Publisher
echo pop-up hiding behind this window. If you see a
echo "Security Notice", click Disable - not Enable.
echo.
echo Press any key to start converting...
pause >nul
echo.
echo Converting Publisher files to PDF...
echo Please wait. This may take a while.
echo.

rem Work from the folder this file is in (pushd also copes with network paths)
pushd "%~dp0"
set "BATFILE=%~f0"
set "PUBVERSION=v7"

rem Every run gets its own log, report and to-do list, named with the date and time, so nothing is overwritten
set "LOGDIR=%~dp0PublisherToPDF-logs"
if not exist "%LOGDIR%" mkdir "%LOGDIR%"
for /f "delims=" %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HH-mm-ss"') do set "STAMP=%%i"
if not defined STAMP set "STAMP=run"
set "LOGFILE=%LOGDIR%\PublisherToPDF-log-%STAMP%.txt"
set "REPORTFILE=%LOGDIR%\PublisherToPDF-report-%STAMP%.txt"
set "TODOFILE=%LOGDIR%\PublisherToPDF-todo-%STAMP%.txt"

> "%LOGFILE%" echo PublishertoPDF %PUBVERSION% - run started %date% %time%

rem If the next line is still in the report afterwards, the run did not finish.
> "%REPORTFILE%" echo This run did not finish normally, so there is no summary. See the log with the same date and time for what was done up to that point.

rem PowerShell reads its own script from the bottom of this same file
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText($env:BATFILE); $i=$t.LastIndexOf('#'+'#PS-SCRIPT-BEGIN'); & ([scriptblock]::Create($t.Substring($i))) -Filter '*.pub' -Recurse"
set "RC=%errorlevel%"

popd

echo.
echo ============================================
echo                 FINISHED
echo ============================================
echo.
echo If it worked, the PDF files are now next to the
echo original Publisher files.
echo.
echo A detailed log and a short report were saved in the
echo folder PublisherToPDF-logs next to this file:
echo "%REPORTFILE%"
echo.
if not exist "%TODOFILE%" goto notodo
echo Some files still need doing. Their names are in:
echo "%TODOFILE%"
echo To do only those, drag that file onto this .bat file.
echo.
:notodo
if "%RC%"=="0" goto closing
echo Something did not go to plan. The report will now open
echo in Notepad. Please send that file to whoever set
echo this up for you.
echo Tip: in Notepad press Ctrl+A then Ctrl+C to copy
echo everything, then paste it into your message.
start "" notepad "%REPORTFILE%"
echo.

:closing
echo Press any key to close this window.
pause >nul

endlocal
exit /b
##PS-SCRIPT-BEGIN
<#
.SYNOPSIS
Converts Microsoft Publisher .pub files to PDF format.

.DESCRIPTION
This script automates the conversion of Microsoft Publisher (.pub) files to PDF format using Publisher's COM automation.
It processes all files matching the specified filter, skips any file that already has a PDF next to it,
and reports successful conversions and any errors encountered during the process.
Everything is also written to the log file named in the LOGFILE environment variable.
Each file is opened read-only in its own fresh Publisher instance, which is closed and released afterwards,
so one bad file cannot upset the ones after it.
If Publisher's normal (typed) automation interface cannot be loaded on this PC, it switches to
late-bound calls, which do not need that interface.

.PARAMETER Filter
Specifies the file filter to select Publisher files for conversion.
This can be a specific file name (e.g., "document.pub") or a wildcard pattern (e.g., "*.pub").

.PARAMETER Recurse
If specified, searches for Publisher files recursively in all subdirectories that match the filter. If omitted, only the current directory is searched.
#>
param
(
    [ValidateNotNullOrEmpty()]
    [string]
    $Filter,

    [switch]
    $Recurse
)

$LogFile = $env:LOGFILE;
$ReportFile = $env:REPORTFILE;
$TodoFile = $env:TODOFILE;
$ListFile = $env:LISTFILE;

function Write-Log {
    param(
        [string]$Message,
        [string]$Color = 'Gray'
    )
    Write-Host $Message -ForegroundColor $Color;
    if ($LogFile) {
        try {
            Add-Content -LiteralPath $LogFile -Value ((Get-Date -Format 'HH:mm:ss') + '  ' + $Message) -Encoding UTF8;
        } catch { }
    }
}

function Test-TypeLibError {
    param([string]$Message)
    return ($Message -match '0x80029C4A|TYPE_E_CANTLOADLIBRARY|Unable to cast COM object');
}

# Late-bound COM call: goes through IDispatch and avoids Publisher's typed interface
function Invoke-Late {
    param(
        $Target,
        [string]$Name,
        [object[]]$Arguments
    )
    return [System.__ComObject].InvokeMember($Name, [System.Reflection.BindingFlags]::InvokeMethod, $null, $Target, $Arguments);
}

# Remember (and log once) that the typed interface is broken, so late-bound calls are used from now on
function Enable-LateBound {
    param([string]$Err)
    if (-not $state.LateBound) {
        $state.LateBound = $true;
        Write-Log "Publisher's normal automation interface is not working on this PC. Switching to a different method." 'Yellow';
        Write-Log ("Details: " + $Err) 'DarkGray';
    }
}

# Release a COM object so Publisher is not kept alive by our reference to it
function Release-Com {
    param($Obj)
    if ($Obj) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) | Out-Null; } catch { }
    }
}

# Ask Publisher to open files with macros and other active content disabled, without prompting
# (3 = msoAutomationSecurityForceDisable). Returns $true if Publisher accepted the setting.
function Set-PublisherSecurity {
    if (-not $state.LateBound) {
        try {
            $app.AutomationSecurity = 3;
            return $true;
        } catch {
            $err = $_.Exception.Message;
            if (Test-TypeLibError $err) {
                Enable-LateBound $err;
            }
        }
    }
    try {
        [System.__ComObject].InvokeMember('AutomationSecurity', [System.Reflection.BindingFlags]::SetProperty, $null, $app, @(3)) | Out-Null;
        return $true;
    } catch {
        return $false;
    }
}

# Opens a document READ-ONLY (so Publisher can never save changes to the original) and without adding it to
# Publisher's recent files list: Open(FileName, ReadOnly = $true, AddToRecentFiles = $false).
# Tries the normal way first and switches to late-bound calls if that interface is broken.
function Open-Doc {
    param([string]$Path)
    if (-not $state.LateBound) {
        try {
            return $app.Open($Path, $true, $false);
        } catch {
            $err = $_.Exception.Message;
            if (Test-TypeLibError $err) {
                Enable-LateBound $err;
            } else {
                throw;
            }
        }
    }
    return (Invoke-Late $app 'Open' @($Path, $true, $false));
}

if (-not $PSBoundParameters.ContainsKey('Filter')) {
    Write-Log "The -Filter parameter is required." 'Red';
    exit 1;
}

if (-not ($Filter -like "*.pub")) {
    Write-Log "The filter must specify .pub files (e.g., '*.pub' or 'file.pub')." 'Red';
    exit 1;
}

# Diagnostic details that go at the top of the log
try {
    Write-Log ("Publisher to PDF conversion started (version " + $env:PUBVERSION + ").");
    Write-Log ("Windows: " + [Environment]::OSVersion.VersionString);
    Write-Log ("PowerShell version " + $PSVersionTable.PSVersion + ", 64-bit process: " + [Environment]::Is64BitProcess);
    Write-Log ("Folder: " + (Get-Location).Path);
    foreach ($root in @($env:ProgramW6432, ${env:ProgramFiles(x86)})) {
        if ($root) {
            foreach ($sub in @('Microsoft Office\root\Office16', 'Microsoft Office\Office16', 'Microsoft Office\root\Office15', 'Microsoft Office\Office15', 'Microsoft Office\Office14')) {
                $p = Join-Path $root ($sub + '\MSPUB.EXE');
                if (Test-Path -LiteralPath $p) {
                    Write-Log ("Publisher found: " + $p + " (version " + (Get-Item -LiteralPath $p).VersionInfo.ProductVersion + ")");
                }
            }
        }
    }
} catch { }

$app = $null;
$doc = $null;
$exitCode = 0;
$state = @{ LateBound = $false; SecurityLogged = $false };
$failedList = New-Object System.Collections.ArrayList;

try {
    if ($ListFile) {
        # List mode: only convert the files named in the list (one full path per line)
        Write-Log ("List mode: only the files named in " + $ListFile + " will be converted.");
        $files = @();
        foreach ($entry in (Get-Content -LiteralPath $ListFile -Encoding UTF8)) {
            $listedPath = $entry.Trim().Trim('"');
            if (-not $listedPath) { continue; }
            if (($listedPath -like "*.pub") -and (Test-Path -LiteralPath $listedPath)) {
                $files += Get-Item -LiteralPath $listedPath;
            } else {
                Write-Log ("Ignored (not a .pub file, or not found): " + $listedPath) 'Yellow';
            }
        }
    } else {
        $files = Get-ChildItem $Filter -File -Recurse:$Recurse -ErrorAction SilentlyContinue;
    }
    if (-not $files) {
        if ($ListFile) {
            Write-Log "None of the entries in the list could be used." 'Yellow';
        } else {
            Write-Log "No Publisher (.pub) files were found in this folder or its subfolders." 'Yellow';
        }
        exit 0;
    }

    Write-Log ("Found " + @($files).Count + " Publisher file(s).");
    Write-Log "If it seems stuck, look for a Publisher pop-up hiding behind this window (click Disable on any Security Notice)." 'Yellow';

    if (-not ([System.Type]::GetTypeFromProgID('Publisher.Application'))) {
        Write-Log "Microsoft Publisher does not appear to be installed on this PC." 'Red';
        exit 1;
    }

    $successCount = 0;
    $skipCount = 0;
    $failCount = 0;
    $consecutiveFails = 0;

    foreach ($file in $files) {
        if ($consecutiveFails -ge 5) {
            Write-Log "Stopping early: 5 files in a row failed, so something is wrong. Please send this log file." 'Red';
            $exitCode = 1;
            break;
        }

        if ($file.Extension -eq ".pub") {
            $fileFullName = $file.FullName;
            $pdfFilePath = [System.IO.Path]::ChangeExtension($fileFullName, '.pdf');
            if (Test-Path -LiteralPath $pdfFilePath) {
                Write-Log "Skipped (PDF already exists): $pdfFilePath";
                $skipCount++;
                Continue;
            }

            Write-Log "Opening: $fileFullName";
            $ok = $false;
            $app = $null;
            $doc = $null;
            try {
                # A fresh Publisher for every file, so one bad file cannot upset the next one
                $app = New-Object -ComObject Publisher.Application;

                $secOk = Set-PublisherSecurity;
                if (-not $state.SecurityLogged) {
                    $state.SecurityLogged = $true;
                    if ($secOk) {
                        Write-Log "Publisher set to open files with active content disabled, without prompts.";
                    } else {
                        Write-Log "Could not switch off Publisher's security pop-ups; click Disable if one appears." 'Yellow';
                    }
                }

                $doc = Open-Doc $fileFullName;
                if (-not $doc) {
                    throw "Publisher did not open the file.";
                }
                Write-Log "  opened, exporting..." 'DarkGray';

                # Export file as PDF (2 = pbFixedFormatTypePDF)
                if ($state.LateBound) {
                    Invoke-Late $doc 'ExportAsFixedFormat' @(2, $pdfFilePath) | Out-Null;
                } else {
                    $doc.ExportAsFixedFormat(2, $pdfFilePath);
                }
                if (Test-Path -LiteralPath $pdfFilePath) {
                    $ok = $true;
                } else {
                    throw "The export finished but no PDF was created.";
                }
            } catch {
                $failCount++;
                $consecutiveFails++;
                [void]$failedList.Add($fileFullName + " -- " + $_.Exception.Message);
                Write-Log ("ERROR converting " + $fileFullName + " -- " + $_.Exception.Message) 'Red';
            } finally {
                # Close the document and quit Publisher, then release both and collect garbage
                if ($doc) {
                    try {
                        if ($state.LateBound) { Invoke-Late $doc 'Close' $null | Out-Null; } else { $doc.Close(); }
                    } catch { }
                }
                if ($app) {
                    try {
                        if ($state.LateBound) { Invoke-Late $app 'Quit' $null | Out-Null; } else { $app.Quit(); }
                    } catch {
                        try { Invoke-Late $app 'Quit' $null | Out-Null; } catch { }
                    }
                }
                Release-Com $doc;
                Release-Com $app;
                $doc = $null;
                $app = $null;
                [System.GC]::Collect();
                [System.GC]::WaitForPendingFinalizers();
            }

            if ($ok) {
                Write-Log "Exported: $pdfFilePath" 'Green';
                $successCount++;
                $consecutiveFails = 0;
            }

            # Give Publisher a moment to finish closing before the next file (up to about 5 seconds)
            $waits = 0;
            while ((Get-Process -Name 'MSPUB' -ErrorAction SilentlyContinue) -and ($waits -lt 10)) {
                Start-Sleep -Milliseconds 500;
                $waits++;
            }
        }
    }

    $summaryColor = 'Green';
    if ($failCount -gt 0) {
        $summaryColor = 'Red';
        $exitCode = 1;
    }
    if ($state.LateBound) {
        Write-Log "Note: the different (late-bound) method was used for this run.";
    }
    Write-Log ("Converted " + $successCount + " files, skipped " + $skipCount + " (PDF already existed), " + $failCount + " errors.") $summaryColor;

    # Work out exactly what is still without a PDF, checked against the disk (not just our own counts).
    # This is wrapped in its own try/catch so a problem in the reporting can never affect the conversions.
    try {
        $remaining = New-Object System.Collections.ArrayList;
        $emptyPdfs = New-Object System.Collections.ArrayList;
        $totalPub = 0;
        foreach ($f in $files) {
            if ($f.Extension -eq ".pub") {
                $totalPub++;
                $pdfCheck = [System.IO.Path]::ChangeExtension($f.FullName, '.pdf');
                if (Test-Path -LiteralPath $pdfCheck) {
                    if ((Get-Item -LiteralPath $pdfCheck).Length -eq 0) { [void]$emptyPdfs.Add($pdfCheck); }
                } else {
                    [void]$remaining.Add($f.FullName);
                }
            }
        }
        $withPdf = $totalPub - $remaining.Count;
        $notTried = $remaining.Count - $failCount;
        if ($notTried -lt 0) { $notTried = 0; }
        $remColor = 'Green';
        if (($remaining.Count -gt 0) -or ($emptyPdfs.Count -gt 0)) {
            $remColor = 'Red';
            $exitCode = 1;
        }

        $summary = @(
            ("Publisher files found in total:      " + $totalPub),
            ("Converted during this run:           " + $successCount),
            ("Skipped (already had a PDF):         " + $skipCount),
            ("Failed during this run:              " + $failCount),
            ("Not attempted (run stopped early):   " + $notTried),
            ("Now have a PDF next to them:         " + $withPdf + " of " + $totalPub),
            ("Still WITHOUT a PDF:                 " + $remaining.Count),
            ("Empty (0 KB) PDFs to delete + redo:  " + $emptyPdfs.Count)
        );
        Write-Log "================ SUMMARY ================";
        foreach ($line in $summary) { Write-Log $line $remColor; }
        if ($failedList.Count -gt 0) {
            Write-Log "Files that failed during this run:" 'Red';
            foreach ($x in $failedList) { Write-Log ("  " + $x) 'Red'; }
        }
        if ($emptyPdfs.Count -gt 0) {
            Write-Log "Empty (0 KB) PDFs found - delete these and run again (they are listed in the report)." 'Red';
        }

        if ($ReportFile) {
            $out = New-Object System.Collections.ArrayList;
            [void]$out.Add("Publisher to PDF report (version " + $env:PUBVERSION + ") - finished " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'));
            [void]$out.Add("Folder: " + (Get-Location).Path);
            if ($ListFile) {
                [void]$out.Add("Mode: only the files named in " + $ListFile);
            } else {
                [void]$out.Add("Mode: whole folder (files that already have a PDF are skipped)");
            }
            [void]$out.Add("");
            foreach ($line in $summary) { [void]$out.Add($line); }
            [void]$out.Add("");
            if ($failedList.Count -gt 0) {
                [void]$out.Add("FAILED DURING THIS RUN (" + $failedList.Count + "), with the reason:");
                foreach ($x in $failedList) { [void]$out.Add($x); }
                [void]$out.Add("");
            }
            if ($remaining.Count -gt 0) {
                [void]$out.Add("STILL WITHOUT A PDF (" + $remaining.Count + ") - these are the ones still to do:");
                foreach ($x in $remaining) { [void]$out.Add($x); }
                if ($TodoFile) {
                    Set-Content -LiteralPath $TodoFile -Value $remaining -Encoding UTF8;
                    Write-Log ("A plain list of the files still to do was saved as: " + $TodoFile);
                    [void]$out.Add("");
                    [void]$out.Add("A plain list of these files was saved as: " + $TodoFile);
                    [void]$out.Add("To redo only these, drag that list file onto the PublishertoPDF .bat file.");
                }
            }
            if ($emptyPdfs.Count -gt 0) {
                [void]$out.Add("");
                [void]$out.Add("EMPTY (0 KB) PDFs (" + $emptyPdfs.Count + ") - delete these, then run again so they are redone:");
                foreach ($x in $emptyPdfs) { [void]$out.Add($x); }
            }
            if (($remaining.Count -eq 0) -and ($emptyPdfs.Count -eq 0)) {
                [void]$out.Add("Every Publisher file has a PDF next to it. Nothing is left to do.");
            }
            Set-Content -LiteralPath $ReportFile -Value $out -Encoding UTF8;
        }
    } catch {
        Write-Log ("Could not write the summary/report: " + $_.Exception.Message) 'Yellow';
    }
} catch {
    Write-Log ("Unexpected problem: " + $_.Exception.Message) 'Red';
    $exitCode = 1;
} finally {
    # Safety net: make sure no Publisher is left running if something went wrong mid-file
    if ($app) {
        try {
            if ($state.LateBound) { Invoke-Late $app 'Quit' $null | Out-Null; } else { $app.Quit(); }
        } catch { }
    }
}

exit $exitCode;
