# References & sources used to create script:
# [https://passbe.com/2019/08/01/bulk-export-onenote-2013-2016-pages-as-html/](https://passbe.com/2019/08/01/bulk-export-onenote-2013-2016-pages-as-html/)
# [https://stackoverflow.com/questions/53689087/powershell-and-onenote](https://stackoverflow.com/questions/53689087/powershell-and-onenote)

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
    $timeoutSeconds = 20
    $startTime = Get-Date
    
    Log-Error "Starting wait loop for page '$pageName'" "DEBUG"
    
    try {
        [void]$onenote.NavigateTo($pageID)
    } catch {
        Log-Error "Failed to NavigateTo page '$pageName'. Error: $_"
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
            Log-Error "Failed to GetPageContent for page '$pageName'. Error: $_"
            return $false
        }
        
        $hasTextPlaceholder = ($xml -match "Wait for OneNote") -or ($xml -match "Wait for onenote")
        $xmlDoc = [xml]$xml
    
        $pendingImages = $xmlDoc | Select-Xml -XPath "//one:Image[not(one:Data) and not(@pathCache) and not(one:CallbackID) and not(@backgroundImage='true')]" -Namespace $schema
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
        }
        
        Start-Sleep -Seconds 1
    } while ((Get-Date) -lt $startTime.AddSeconds($timeoutSeconds))
    
    if ($isWaiting) { Write-Host "" }
    $errMsg = "TIMEOUT: Page '$pageName' failed to download all assets. Exporting incomplete page."
    Write-Host "      -> [ERROR] $errMsg" -ForegroundColor Red
    Log-Error $errMsg
    return $false
}

# --- EXTRACT AND ALIGN MISSING IMAGES (Printouts, Lens Scans, etc) ---
Function Extract-Callback-Images {
    param( $onenote, $xml, $pageID, $htmlFilePath, $attachmentsPath, $pageName )
    try {
        $schema = @{one="http://schemas.microsoft.com/office/onenote/2013/onenote"}
        $xmlDoc = [xml]$xml
        
        # Only look for images outside the Title area
        $callbackImages = $xmlDoc | Select-Xml -XPath "//one:Page/one:Outline//one:Image[one:CallbackID] | //one:Page/one:Image[one:CallbackID]" -Namespace $schema
        if (-not $callbackImages) { return }

        $htmlContent = Get-Content -Path $htmlFilePath -Raw
        
        # Wrap injected absolute images in a div that stays behind text and doesn't block mouse selection
        $appendedTags = "`n<!-- Injected Missing Images -->`n<div style='position:absolute; top:0; left:0; width:100%; height:100%; pointer-events:none; z-index:-1;'>`n"

        foreach ($img in $callbackImages) {
            $cbNode = $img.Node | Select-Xml -XPath "one:CallbackID" -Namespace $schema
            if (-not $cbNode) { continue }
            $callbackID = $cbNode.Node.GetAttribute("callbackID")
            if (-not $callbackID) { continue }
            
            $base64String = ""
            try {
                [void]$onenote.GetBinaryPageContent($pageID, $callbackID, [ref]$base64String)
            } catch { continue }
            
            if ($base64String) {
                if (-not (Test-Path $attachmentsPath)) { [void](New-Item -Path $attachmentsPath -ItemType Directory -ErrorAction Ignore) }
                
                $imageBytes = [Convert]::FromBase64String($base64String)
                $imageFileName = "img_$($callbackID -replace '\{|\}|-','').png"
                $imagePath = Join-Path -Path $attachmentsPath -ChildPath $imageFileName
                [System.IO.File]::WriteAllBytes($imagePath, $imageBytes)
                
                $folderName = Split-Path $attachmentsPath -Leaf
                $relativeImgUrl = "$folderName/$imageFileName"
                
                # --- CALCULATE EXACT POSITION AND SCALE ---
                $x = 0.0
                $y = 0.0
                $width = "auto"
                $height = "auto"
                
                $sizeNode = $img.Node | Select-Xml -XPath "one:Size" -Namespace $schema
                if ($sizeNode) {
                    $w = $sizeNode.Node.GetAttribute("width")
                    $h = $sizeNode.Node.GetAttribute("height")
                    if ($w) { $width = "$w`pt" }
                    if ($h) { $height = "$h`pt" }
                }
                
                # Traverse up the XML tree to find the absolute X/Y of the parent Outline
                $curr = $img.Node
                while ($curr -ne $null -and $curr.Name -ne "one:Page") {
                    $posNode = $curr | Select-Xml -XPath "one:Position" -Namespace $schema
                    if ($posNode) {
                        $px = $posNode.Node.GetAttribute("x")
                        $py = $posNode.Node.GetAttribute("y")
                        if ($px) { $x += [double]$px }
                        if ($py) { $y += [double]$py }
                    }
                    $curr = $curr.ParentNode
                }
                
                # Apply absolute CSS styling
                $style = "position: absolute; left: ${x}pt; top: ${y}pt; "
                if ($width -ne "auto") { $style += "width: $width; " }
                if ($height -ne "auto") { $style += "height: $height; " }
                $style += "mix-blend-mode: multiply;"
                
                $appendedTags += "<img src='$relativeImgUrl' style='$style' alt='Extracted Callback Image' />`n"
                Write-Host "      -> [Media] Restored hidden image at x:${x}pt, y:${y}pt" -ForegroundColor Cyan
            }
        }
        $appendedTags += "</div>`n"
        
        if ($appendedTags -match "<img") {
            $htmlContent = $htmlContent -replace "(?i)(</body>)", "$appendedTags`n`$1"
            Set-Content -Path $htmlFilePath -Value $htmlContent -Encoding UTF8
        }
    } catch {
        Log-Error "Failed to extract Callback images for page '$pageName'. Error: $_"
    }
}

# --- EXTRACT HANDWRITTEN TITLE ---
Function Extract-Ink-Title {
    param ( $onenote, $xml, $pageID, $htmlFilePath, $attachmentsPath, $pageName )
    try {
        $schema = @{one="http://schemas.microsoft.com/office/onenote/2013/onenote"}
        $xmlDoc = [xml]$xml
        
        # FIX: Directly target the CallbackID inside the Title block regardless of whether it's an InkWord or InkDrawing
        $callbackNode = $xmlDoc | Select-Xml -XPath "//one:Page/one:Title//one:CallbackID" -Namespace $schema
        
        if ($callbackNode) {
            $callbackID = $callbackNode.Node.GetAttribute("callbackID")
            if (-not $callbackID) { return $null }

            $base64String = ""
            try {
                [void]$onenote.GetBinaryPageContent($pageID, $callbackID, [ref]$base64String)
            } catch { return $null }
            
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
    } catch { }
    return $null
}

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
            $spacingPts = if ($horizontal -and $horizontal.Node.spacing) { [double]$horizontal.Node.spacing } else { 23.76 }
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
    } catch { }
}

Function Export-OneNote-Page {
    param( $onenote, $node, $path )
    $name = ReplaceIllegal -text $node.name
    $file = $(Join-Path -Path $path -ChildPath "$($name).htm")
    Write-Host "    Page: $($file)"
    
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
    
    # Process standard file attachments
    try {
        $schema = @{one="http://schemas.microsoft.com/office/onenote/2013/onenote"}
        $xmlDoc = [xml]$xml
        $xmlDoc | Select-Xml -XPath "//one:Page/one:Outline/one:OEChildren/one:OE/one:InsertedFile" -Namespace $schema | foreach {
            $attFile = Join-Path -Path $attachmentpath -ChildPath $_.Node.preferredName
            if ($_.Node.pathCache) {
                [void](Copy-Item $_.Node.pathCache -Destination $attFile -ErrorAction SilentlyContinue)
            }
        }
    } catch { }

    # Extract missing printouts/scans and align them properly
    Extract-Callback-Images -onenote $onenote -xml $xml -pageID $node.ID -htmlFilePath $file -attachmentsPath $attachmentpath -pageName $name

    $result = New-Object psobject -Property @{
        FilePath = $file
        InkTitleUrl = $inkTitleUrl
    }
    return $result
}

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
        } catch { }
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
    Write-Host "CRITICAL ERROR: Failed to connect to OneNote." -ForegroundColor Red
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
    } catch { }
}

Get-ChildItem -path $folder filelist.xml -Recurse | foreach { [void](Remove-Item -Path $_.FullName -ErrorAction Ignore) }

$finishTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Out-File -FilePath $global:errorLogPath -InputObject "=== OneNote Export Finished: $finishTime ===" -Append -Encoding UTF8
Write-Host "`nExport Complete!" -ForegroundColor Cyan
