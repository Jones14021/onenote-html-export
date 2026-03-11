# References & sources used to create script:
# [https://passbe.com/2019/08/01/bulk-export-onenote-2013-2016-pages-as-html/](https://passbe.com/2019/08/01/bulk-export-onenote-2013-2016-pages-as-html/)
# [https://stackoverflow.com/questions/53689087/powershell-and-onenote](https://stackoverflow.com/questions/53689087/powershell-and-onenote)
# [http://thebackend.info/powershell/2017/12/onenote-read-and-write-content-with-powershell/](http://thebackend.info/powershell/2017/12/onenote-read-and-write-content-with-powershell/)

# --- Global Configuration & Logging ---
$global:debugMode = $true

Function Log-Error {
    param( [string]$message, [string]$type = "ERROR" )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] [$type] $message"
    Out-File -FilePath $global:errorLogPath -InputObject $logLine -Append -Encoding UTF8
    if ($global:debugMode) { Write-Host $logLine -ForegroundColor DarkGray }
}

# --- Actively wait for OneNote to finish downloading lazy-loaded content ---
Function Wait-For-Page-Load {
    param( $onenote, $pageID, $pageName )
    $timeoutSeconds = 20 # Increased slightly for debugging
    $startTime = Get-Date
    
    Log-Error "Starting wait loop for page '$pageName'" "DEBUG"
    
    try {
        [void]$onenote.NavigateTo($pageID)
    } catch {
        Log-Error "Failed to NavigateTo page '$pageName' (ID: $pageID). Error: $_"
        return $false
    }
    
    $schema = @{one="http://schemas.microsoft.com/office/onenote/2013/onenote"}
    $lastPrintedSecond = -1
    $isWaiting = $false
    
    do {
        $xml = ""
        try {
            [void]$onenote.GetPageContent($pageID, [ref]$xml, 0)
        } catch {
            if ($isWaiting) { Write-Host "" } 
            Log-Error "Failed to GetPageContent for page '$pageName' during wait loop. Error: $_"
            return $false
        }
        
        $hasTextPlaceholder = ($xml -match "Wait for OneNote") -or ($xml -match "Wait for onenote")
        $xmlDoc = [xml]$xml
    
        $pendingImages = $xmlDoc | Select-Xml -XPath "//one:Image[not(one:Data) and not(@pathCache) and not(@backgroundImage='true')]" -Namespace $schema
        $pendingFiles = $xmlDoc | Select-Xml -XPath "//one:InsertedFile[not(@pathCache)]" -Namespace $schema
        
        $hasBackgroundImages = $null -ne ($xmlDoc | Select-Xml -XPath "//one:Image[@backgroundImage='true']" -Namespace $schema)
        $hasAnyForegroundMedia = $null -ne ($xmlDoc | Select-Xml -XPath "//one:Image[not(@backgroundImage='true')] | //one:InsertedFile" -Namespace $schema)
        
        $onlyBackgroundsExist = ($hasBackgroundImages -and -not $hasAnyForegroundMedia)

        if (-not $hasTextPlaceholder -and ($onlyBackgroundsExist -or (($null -eq $pendingImages) -and ($null -eq $pendingFiles)))) {
            Start-Sleep -Milliseconds 300
            if ($isWaiting) { Write-Host "" }
            Log-Error "Page '$pageName' loaded successfully." "DEBUG"
            return $true
        }
        
        $elapsedSeconds = [math]::Floor(((Get-Date) - $startTime).TotalSeconds)
        
        if ($elapsedSeconds -ne $lastPrintedSecond) {
            $msg = "      -> [Wait] Waiting for '$pageName' to load... ($elapsedSeconds of $timeoutSeconds seconds)"
            Write-Host "`r$($msg.PadRight(90, ' '))" -NoNewline -ForegroundColor DarkCyan
            $lastPrintedSecond = $elapsedSeconds
            $isWaiting = $true
            
            # --- INSTRUMENTATION: Dump missing image IDs to log ---
            if ($pendingImages) {
                foreach ($img in $pendingImages) {
                    Log-Error "Pending Image found on '$pageName'. ObjID: $($img.Node.objectID)" "DEBUG-IMG"
                }
            }
        }
        
        Start-Sleep -Seconds 1
    } while ((Get-Date) -lt $startTime.AddSeconds($timeoutSeconds))
    
    if ($isWaiting) { Write-Host "" }
    $errMsg = "TIMEOUT: Page '$pageName' failed to download all assets. Exporting incomplete page."
    Write-Host "      -> [ERROR] $errMsg" -ForegroundColor Red
    Log-Error $errMsg
    
    # --- INSTRUMENTATION: Dump the XML of the failed page ---
    Log-Error "DUMPING FAILED XML FOR '$pageName': `n$xml" "XML-DUMP"
    
    return $false
}


# --- Extract handwritten ink titles as PNG images ---
Function Extract-Ink-Title {
    param ( $onenote, $xml, $pageID, $htmlFilePath, $attachmentsPath, $pageName )
    try {
        $schema = @{one="http://schemas.microsoft.com/office/onenote/2013/onenote"}
        $xmlDoc = [xml]$xml
        $inkTitleNode = $xmlDoc | Select-Xml -XPath "//one:Page/one:Title//one:InkWord" -Namespace $schema
        
        if ($inkTitleNode -and $inkTitleNode.Node.CallbackID) {
            $callbackID = $inkTitleNode.Node.CallbackID
            $base64String = ""
            [void]$onenote.GetBinaryPageContent($pageID, $callbackID, [ref]$base64String)
            
            if ($base64String) {
                if (-not (Test-Path $attachmentsPath)) { [void](New-Item -Path $attachmentsPath -ItemType Directory -ErrorAction Ignore) }
                
                $imageBytes = [Convert]::FromBase64String($base64String)
                $imageFileName = "InkTitle_$($pageID -replace '\{|\}|-','').png"
                $imagePath = Join-Path -Path $attachmentsPath -ChildPath $imageFileName
                [System.IO.File]::WriteAllBytes($imagePath, $imageBytes)
                
                $folderName = Split-Path $attachmentsPath -Leaf
                $relativeImgUrl = "$folderName/$imageFileName"
                
                $htmlContent = Get-Content -Path $htmlFilePath -Raw
                $imgTag = "<div style='padding: 20px 0px;'><img src='$relativeImgUrl' style='max-height: 80px; mix-blend-mode: multiply;' alt='Handwritten Title' /></div>"
                $htmlContent = $htmlContent -replace "(?i)(<body[^>]*>)", "`$1`n$imgTag"
                Set-Content -Path $htmlFilePath -Value $htmlContent -Encoding UTF8
                
                return $relativeImgUrl
            }
        }
    } catch {
        Log-Error "Failed to extract Ink Title for page '$pageName'. Error: $_"
    }
    return $null
}

# --- Dynamically recreate OneNote rule/grid lines in HTML via CSS ---
Function Inject-HTML-Background {
    param ( $xml, $htmlFilePath, $pageName )
    try {
        $schema = @{one="http://schemas.microsoft.com/office/onenote/2013/onenote"}
        $xmlDoc = [xml]$xml
        $ruleLines = $xmlDoc | Select-Xml -XPath "//one:RuleLines" -Namespace $schema
        
        $baseCss = "<style> div > img { mix-blend-mode: multiply; } "
        
        if ($ruleLines -and $ruleLines.Node.visible -eq "true") {
            $isGrid = $null -ne ($ruleLines.Node | Select-Xml -XPath "one:Vertical" -Namespace $schema)
            $horizontal = $ruleLines.Node | Select-Xml -XPath "one:Horizontal" -Namespace $schema
            
            $spacingPts = 23.76
            if ($horizontal -and $horizontal.Node.spacing) { $spacingPts = [double]$horizontal.Node.spacing }
            $spacingPx = [math]::Round($spacingPts * 1.33)
            $lineColor = "#d1e1e8" 
            
            if ($isGrid) {
                $baseCss += "body { background-color: white !important; background-size: ${spacingPx}px ${spacingPx}px !important; background-image: linear-gradient(to right, $lineColor 1px, transparent 1px), linear-gradient(to bottom, $lineColor 1px, transparent 1px) !important; }"
            } else {
                $baseCss += "body { background-color: white !important; background-size: 100% ${spacingPx}px !important; background-image: linear-gradient(transparent $([math]::Max(1, $spacingPx - 1))px, $lineColor 1px) !important; }"
            }
        }
        $baseCss += "</style>"
        $htmlContent = Get-Content -Path $htmlFilePath -Raw
        $htmlContent = $htmlContent -replace "(?i)</head>", "`n$baseCss`n</head>"
        Set-Content -Path $htmlFilePath -Value $htmlContent -Encoding UTF8
    } catch {
        Log-Error "Failed to inject CSS rule lines for page '$pageName'. Error: $_"
    }
}

# --- Export page ---
Function Export-OneNote-Page {
    param( $onenote, $node, $path )
    $name = ReplaceIllegal -text $node.name
    $file = $(Join-Path -Path $path -ChildPath "$($name).htm")
    Write-Host "    Page: $($file)"
    
    # Cast entirely to void to protect pipeline
    [void](Wait-For-Page-Load -onenote $onenote -pageID $node.ID -pageName $name)

    try {
        [void]$onenote.Publish($node.ID, $file, 7, "")
        $xml = ''
        [void]$onenote.GetPageContent($node.ID, [ref]$xml, 0)
        Inject-HTML-Background -xml $xml -htmlFilePath $file -pageName $name
    } catch {
        Log-Error "Publish() failed for page '$name'. Error: $_"
        return $null
    }
    
    $attachmentpath = Join-Path -Path $path -ChildPath ($name + "_files")
    $inkTitleUrl = $null
    if ($name -match "Untitled Page" -or $name -match "Unbenannte Seite") {
        $inkTitleUrl = Extract-Ink-Title -onenote $onenote -xml $xml -pageID $node.ID -htmlFilePath $file -attachmentsPath $attachmentpath -pageName $name
    }
    
    # --- INSTRUMENTATION: Modified Attachment Export ---
    try {
        $schema = @{one="http://schemas.microsoft.com/office/onenote/2013/onenote"}
        $xmlDoc = [xml]$xml
        $xmlDoc | Select-Xml -XPath "//one:Page/one:Outline/one:OEChildren/one:OE/one:InsertedFile" -Namespace $schema | foreach {
            $attFile = Join-Path -Path $attachmentpath -ChildPath $_.Node.preferredName
            if ($_.Node.pathCache) {
                [void](Copy-Item $_.Node.pathCache -Destination $attFile -ErrorAction SilentlyContinue)
            } else {
                Log-Error "Missing pathCache for attachment '$($_.Node.preferredName)' on page '$name'" "ATT-MISSING"
            }
        }
    } catch { Log-Error "Attachment extraction failed. $_" }

    # Strict return object construction
    $result = New-Object psobject -Property @{
        FilePath = $file
        InkTitleUrl = $inkTitleUrl
    }
    Log-Error "Export-OneNote-Page Returning Object. FilePath: $($result.FilePath)" "DEBUG"
    return $result
}

# --- Spider (Updated to strictly manipulate a StringBuilder instead of returning string pipeline) ---
Function Spider-OneNote-Notebook {
    param( $onenote, $node, $path, $notebookRoot, $htmlBuilder )
    
    $previouslevel = 0
    $previousname = ""
    $grandparent = ""
    $parent = ""

    foreach($child in $node.ChildNodes) {
        try {
            $safeName = ReplaceIllegal -text $child.name
            $levelchange = $child.pageLevel - $previouslevel
            $displayName = [System.Net.WebUtility]::HtmlEncode($child.name)

            if (-not $child.HasChildNodes) {
                if ($levelchange -eq 1) {
                    if ($previouslevel -ne 0) {
                        $grandparent = $parent
                        $parent = $previousname
                    }
                    $filepath = Join-Path -path $(join-path -path $path -ChildPath $grandparent) -ChildPath $parent
                    [void](New-Item -Path $filepath -ItemType directory -ErrorAction Ignore)
                    $pageResult = Export-OneNote-Page -onenote $onenote -node $child -path $filepath
                } elseif ($levelchange -eq -1) {
                    $filepath = Join-Path -path $path -ChildPath $grandparent
                    [void](New-Item -Path $filepath -ItemType directory -ErrorAction Ignore)
                    $pageResult = Export-OneNote-Page -onenote $onenote -node $child -path $filepath
                    $parent = $grandparent
                    $grandparent = ""
                } elseif ($levelchange -eq -2) {
                    $pageResult = Export-OneNote-Page -onenote $onenote -node $child -path $path
                    $parent = ""
                    $grandparent = ""
                } elseif ($levelchange -eq 0 -and $parent -eq "") {
                    $pageResult = Export-OneNote-Page -onenote $onenote -node $child -path $path
                } else {
                    $grandparentpath = Join-Path -path $path -ChildPath $grandparent
                    $filepath = Join-Path -path $grandparentpath -ChildPath $parent
                    [void](New-Item -Path $filepath -ItemType directory -ErrorAction Ignore)
                    $pageResult = Export-OneNote-Page -onenote $onenote -node $child -path $filepath
                }
                
                # --- INSTRUMENTATION: Verify what the Spider received ---
                $typeR = $pageResult.GetType().Name
                Log-Error "Spider received from Export-OneNote-Page. Type: $typeR" "DEBUG-PIPELINE"
                
                if ($null -ne $pageResult -and $null -ne $pageResult.FilePath) {
                    $relPath = $pageResult.FilePath.Substring($notebookRoot.Length + 1)
                    $relUrl = $relPath -replace '\\', '/' -replace ' ', '%20'
                    $indentLevel = [int]$child.pageLevel * 20
                    
                    if ($pageResult.InkTitleUrl) {
                        $parentFolderPath = Split-Path $relUrl -Parent
                        $fullImgUrl = if ($parentFolderPath) { "$parentFolderPath/$($pageResult.InkTitleUrl)" } else { $pageResult.InkTitleUrl }
                        $linkContent = "<img src='$fullImgUrl' style='height: 24px; vertical-align: middle; mix-blend-mode: multiply;' alt='Ink Title' />"
                    } else {
                        $linkContent = $displayName
                    }
                    
                    [void]$htmlBuilder.AppendLine("<li class='page' style='margin-left: $($indentLevel)px;'><a href=`"$relUrl`">$linkContent</a></li>")
                } else {
                    Log-Error "Spider PageResult was null or missing FilePath for '$displayName'" "DEBUG-PIPELINE-FAIL"
                }

            } else {
                $folder = Join-Path -Path $path -ChildPath $safeName
                [void](New-Item -Path $folder -ItemType directory -ErrorAction Ignore)
                Write-Host "  Section: $($folder)"

                [void]$htmlBuilder.AppendLine("<li class='section'>$displayName<ul>")
                Spider-OneNote-Notebook -onenote $onenote -node $child -path $folder -notebookRoot $notebookRoot -htmlBuilder $htmlBuilder
                [void]$htmlBuilder.AppendLine("</ul></li>")
            }

            $previouslevel = $child.pageLevel
            $previousname = $safeName 
        } catch {
            Log-Error "Failed while spidering node. Error: $_" "SPIDER-ERROR"
        }
    }
}

Function ReplaceIllegal {
    param ( $text )
    $illegal = [string]::join('',([System.IO.Path]::GetInvalidFileNameChars())) -replace '\\','\\'
    $replaced = $text -replace "[$illegal]",'_'
    return $replaced
}

Function Get-Folder {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null
    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select an export folder"
    $foldername.rootfolder = "MyComputer"
    if($foldername.ShowDialog() -eq "OK") { return $foldername.SelectedPath }
    return $null
}

# ================= MAIN EXECUTION =================

$folder = Get-Folder
if (-not $folder) { Write-Host "No folder selected. Exiting."; exit }

$global:errorLogPath = Join-Path -Path $folder -ChildPath "errors.log"
Out-File -FilePath $global:errorLogPath -InputObject "=== OneNote Export Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===" -Encoding UTF8

try {
    $OneNote = New-Object -ComObject OneNote.Application
    [xml]$Hierarchy = ""
    [void]$OneNote.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$Hierarchy)
} catch {
    Write-Host "CRITICAL ERROR: Failed to connect to OneNote COM API." -ForegroundColor Red
    Log-Error "Failed to initialize OneNote COM Object. Error: $_"
    exit
}

foreach ($notebook in $Hierarchy.Notebooks.Notebook ) {
    try {
        $name = ReplaceIllegal -text $notebook.name
        $nf = Join-Path -Path $folder -ChildPath $name
        Write-Host "=======================================" -ForegroundColor Magenta
        Write-Host "Notebook: $($nf)" -ForegroundColor Magenta
        Write-Host "=======================================" -ForegroundColor Magenta
        [void](New-Item -Path $nf -ItemType directory -ErrorAction Ignore)
        
        # --- FIXED PIPELINE: Pass a mutable StringBuilder object down the chain ---
        $htmlBuilder = New-Object System.Text.StringBuilder
        Spider-OneNote-Notebook -onenote $OneNote -node $notebook -path $nf -notebookRoot $nf -htmlBuilder $htmlBuilder

        $safeNotebookName = [System.Net.WebUtility]::HtmlEncode($notebook.name)
        $tocBody = $htmlBuilder.ToString()
        
        $indexHtml = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>$safeNotebookName - Index</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; background-color: #f3f2f1; }
        .container { max-width: 900px; margin: 0 auto; background: white; padding: 30px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #7719aa; border-bottom: 2px solid #7719aa; padding-bottom: 10px; }
        ul { list-style-type: none; padding-left: 20px; }
        li { margin: 6px 0; }
        .section { font-size: 1.2em; font-weight: bold; margin-top: 20px; color: #333; }
        .page { font-size: 1em; font-weight: normal; }
        a { text-decoration: none; color: #0078d4; }
        a:hover { text-decoration: underline; color: #004578; }
    </style>
</head>
<body>
    <div class="container">
        <h1>$safeNotebookName</h1>
        <ul>
$tocBody
        </ul>
    </div>
</body>
</html>
"@

        $indexPath = Join-Path -Path $nf -ChildPath "index.htm"
        Set-Content -Path $indexPath -Value $indexHtml -Encoding UTF8
        Write-Host "  -> Created Table of Contents: $($indexPath)"
    } catch {
        Log-Error "Critical error processing notebook '$($notebook.name)'. Error: $_"
    }
}

Get-ChildItem -path $folder filelist.xml -Recurse | foreach { [void](Remove-Item -Path $_.FullName -ErrorAction Ignore) }

$finishTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Out-File -FilePath $global:errorLogPath -InputObject "=== OneNote Export Finished: $finishTime ===" -Append -Encoding UTF8
Write-Host "`nExport Complete! Check '$global:errorLogPath' for debug output and errors." -ForegroundColor Cyan
