# References & sources used to create script:
# [https://passbe.com/2019/08/01/bulk-export-onenote-2013-2016-pages-as-html/](https://passbe.com/2019/08/01/bulk-export-onenote-2013-2016-pages-as-html/)
# [https://stackoverflow.com/questions/53689087/powershell-and-onenote](https://stackoverflow.com/questions/53689087/powershell-and-onenote)
# [http://thebackend.info/powershell/2017/12/onenote-read-and-write-content-with-powershell/](http://thebackend.info/powershell/2017/12/onenote-read-and-write-content-with-powershell/)
# [https://stackoverflow.com/questions/53639041/how-to-access-contents-of-onenote-page](https://stackoverflow.com/questions/53639041/how-to-access-contents-of-onenote-page)

# --- Error Logging Helper ---
Function Log-Error {
    param( [string]$message )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] $message"
    # Write to the script-scoped log file immediately
    Out-File -FilePath $script:errorLogPath -InputObject $logLine -Append -Encoding UTF8
}

# --- Actively wait for OneNote to finish downloading lazy-loaded content ---
Function Wait-For-Page-Load {
    param( $onenote, $pageID, $pageName )
    $timeoutSeconds = 120 # 2 minutes maximum
    $startTime = Get-Date
    
    Write-Host "      -> [Navigation] Triggering download engine..." -ForegroundColor Cyan
    Write-Host "      -> [WARNING] Do not click inside OneNote during this load!" -ForegroundColor Yellow
    
    try {
        $onenote.NavigateTo($pageID)
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
            # 0 = piAll (returns all page content, including binary data and cache paths)
            $onenote.GetPageContent($pageID, [ref]$xml, 0)
        } catch {
            if ($isWaiting) { Write-Host "" } # Break the inline line on error
            Log-Error "Failed to GetPageContent for page '$pageName' during wait loop. Error: $_"
            return $false
        }
        
        # 1. Check if the literal placeholder text exists anywhere on the page
        $hasTextPlaceholder = ($xml -match "Wait for OneNote") -or ($xml -match "Wait for onenote")
        
        $xmlDoc = [xml]$xml
        
        # 2. Explicitly check for OneNote's "CallbackID" tag, which it uses when data is pending download
        $hasCallback = $null -ne ($xmlDoc | Select-Xml -XPath "//*[one:CallbackID]" -Namespace $schema)

        # 3. Parse XML to find empty image nodes. We explicitly IGNORE background images 
        # (not(@backgroundImage='true')) because the API omits their data even when fully loaded.
        $pendingImages = $xmlDoc | Select-Xml -XPath "//one:Image[not(one:Data) and not(@pathCache) and not(@backgroundImage='true')]" -Namespace $schema
        $pendingFiles = $xmlDoc | Select-Xml -XPath "//one:InsertedFile[not(@pathCache)]" -Namespace $schema
        
        # If no placeholders and no pending nodes are found, the page is fully loaded!
        if (-not $hasTextPlaceholder -and -not $hasCallback -and ($null -eq $pendingImages) -and ($null -eq $pendingFiles)) {
            Start-Sleep -Milliseconds 300 # Brief pause to let the OneNote rendering engine catch up
            if ($isWaiting) { Write-Host "" } # End the inline line gracefully
            Write-Host "      -> [Success] Page fully loaded!" -ForegroundColor Green
            return $true
        }
        
        $elapsedSeconds = [math]::Floor(((Get-Date) - $startTime).TotalSeconds)
        
        # Print status update inline using carriage return (`r) and NoNewline
        if ($elapsedSeconds -ne $lastPrintedSecond) {
            $msg = "      -> [Wait] Waiting for '$pageName' to load... ($elapsedSeconds of $timeoutSeconds seconds)"
            # Pad with spaces to ensure it overwrites previous longer lines
            Write-Host "`r$($msg.PadRight(90, ' '))" -NoNewline -ForegroundColor DarkCyan
            $lastPrintedSecond = $elapsedSeconds
            $isWaiting = $true
        }
        
        Start-Sleep -Seconds 1
    } while ((Get-Date) -lt $startTime.AddSeconds($timeoutSeconds))
    
    if ($isWaiting) { Write-Host "" } # End inline line on timeout
    
    # Timeout reached
    $errMsg = "TIMEOUT ERROR: Page '$pageName' (ID: $pageID) failed to download all assets within the $timeoutSeconds second limit."
    Write-Host "      -> [ERROR] $errMsg" -ForegroundColor Red
    Log-Error $errMsg
    return $false
}

# --- Dynamically recreate OneNote rule/grid lines in HTML via CSS ---
Function Inject-HTML-Background {
    param ( $xml, $htmlFilePath, $pageName )
    try {
        $schema = @{one="http://schemas.microsoft.com/office/onenote/2013/onenote"}
        $xmlDoc = [xml]$xml
        $ruleLines = $xmlDoc | Select-Xml -XPath "//one:RuleLines" -Namespace $schema
        
        if ($ruleLines -and $ruleLines.Node.visible -eq "true") {
            $isGrid = $null -ne ($ruleLines.Node | Select-Xml -XPath "one:Vertical" -Namespace $schema)
            $horizontal = $ruleLines.Node | Select-Xml -XPath "one:Horizontal" -Namespace $schema
            
            $spacingPts = 23.76 # OneNote default
            if ($horizontal -and $horizontal.Node.spacing) {
                $spacingPts = [double]$horizontal.Node.spacing
            }
            $spacingPx = [math]::Round($spacingPts * 1.33)
            $lineColor = "#d1e1e8" 
            
            if ($isGrid) {
                Write-Host "      -> [UI] Injecting CSS Grid Lines (${spacingPx}px)" -ForegroundColor Cyan
                $css = "<style> body { background-color: white !important; background-size: ${spacingPx}px ${spacingPx}px !important; background-image: linear-gradient(to right, $lineColor 1px, transparent 1px), linear-gradient(to bottom, $lineColor 1px, transparent 1px) !important; } </style>"
            } else {
                Write-Host "      -> [UI] Injecting CSS Lined Paper (${spacingPx}px)" -ForegroundColor Cyan
                $css = "<style> body { background-color: white !important; background-size: 100% ${spacingPx}px !important; background-image: linear-gradient(transparent $([math]::Max(1, $spacingPx - 1))px, $lineColor 1px) !important; } </style>"
            }
            
            $htmlContent = Get-Content -Path $htmlFilePath -Raw
            $htmlContent = $htmlContent -replace "(?i)</head>", "`n$css`n</head>"
            Set-Content -Path $htmlFilePath -Value $htmlContent -Encoding UTF8
        }
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
    
    Wait-For-Page-Load -onenote $onenote -pageID $node.ID -pageName $name | Out-Null

    try {
        # 1. Export standard HTML
        $onenote.Publish($node.ID, $file, 7, "")
        
        # 2. Get XML once for both Grid Lines and Attachments
        $xml = ''
        $onenote.GetPageContent($node.ID, [ref]$xml, 0)
        
        # 3. Inject CSS Rule/Grid lines
        Inject-HTML-Background -xml $xml -htmlFilePath $file -pageName $name
        
    } catch {
        Log-Error "COM API Publish() or XML retrieval failed for page '$name'. Error: $_"
        return $null
    }
    
    # 4. Export Attachments using already-fetched XML
	$attachmentpath = Join-Path -Path $path -ChildPath ($name + "_files")
    Export-OneNote-Attachments -xml $xml -path $attachmentpath -pageName $name

    return $file
}

# --- Copy embedded attachments (Updated to use passed XML) ---
Function Export-OneNote-Attachments {
    param ( $xml, $path, $pageName )
    try {
        $schema = @{one="http://schemas.microsoft.com/office/onenote/2013/onenote"}
        $xml | Select-Xml -XPath "//one:Page/one:Outline/one:OEChildren/one:OE/one:InsertedFile" -Namespace $schema | foreach {
            $file = Join-Path -Path $path -ChildPath $_.Node.preferredName
            Write-Host "      Attachment: $($file)"
            try {
                Copy-Item $_.Node.pathCache -Destination $file -ErrorAction Stop
            } catch {
                Log-Error "Failed to copy attachment '$($_.Node.preferredName)' for page '$pageName'. PathCache: '$($_.Node.pathCache)'. Error: $_"
            }
        }
    } catch {
        Log-Error "Failed to parse attachments XML for page '$pageName'. Error: $_"
    }
}

Function Spider-OneNote-Notebook {
    param( $onenote, $node, $path, $notebookRoot )
    $tocHtml = ""
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
                # --- It's a Page ---
                if ($levelchange -eq 1) {
                    if ($previouslevel -ne 0) {
                        $grandparent = $parent
                        $parent = $previousname
                    }
                    $filepath = Join-Path -path $(join-path -path $path -ChildPath $grandparent) -ChildPath $parent
                    New-Item -Path $filepath -ItemType directory -ErrorAction Ignore | Out-Null
                    $fileAbsPath = Export-OneNote-Page -onenote $onenote -node $child -path $filepath
                } elseif ($levelchange -eq -1) {
                    $filepath = Join-Path -path $path -ChildPath $grandparent
                    New-Item -Path $filepath -ItemType directory -ErrorAction Ignore | Out-Null
                    $fileAbsPath = Export-OneNote-Page -onenote $onenote -node $child -path $filepath
                    $parent = $grandparent
                    $grandparent = ""
                } elseif ($levelchange -eq -2) {
                    $fileAbsPath = Export-OneNote-Page -onenote $onenote -node $child -path $path
                    $parent = ""
                    $grandparent = ""
                } elseif ($levelchange -eq 0 -and $parent -eq "") {
                    $fileAbsPath = Export-OneNote-Page -onenote $onenote -node $child -path $path
                } else {
                    $grandparentpath = Join-Path -path $path -ChildPath $grandparent
                    $filepath = Join-Path -path $grandparentpath -ChildPath $parent
                    New-Item -Path $filepath -ItemType directory -ErrorAction Ignore | Out-Null
                    $fileAbsPath = Export-OneNote-Page -onenote $onenote -node $child -path $filepath
                }
                
                # Create a relative link for the index.htm
                if ($fileAbsPath) {
                    $relPath = $fileAbsPath.Substring($notebookRoot.Length + 1)
                    $relUrl = $relPath -replace '\\', '/' -replace ' ', '%20'
                    
                    # Use margin-left to visually indent subpages based on their OneNote level
                    $indentLevel = [int]$child.pageLevel * 20
                    $tocHtml += "<li class='page' style='margin-left: $($indentLevel)px;'><a href=`"$relUrl`">$displayName</a></li>`n"
                }
            } else {
                # --- It's a Section or Section Group ---
                $folder = Join-Path -Path $path -ChildPath $safeName
                New-Item -Path $folder -ItemType directory -ErrorAction Ignore | Out-Null
                Write-Host "  Section: $($folder)"

                $tocHtml += "<li class='section'>$displayName<ul>`n"
                # Recursively crawl the section and append its HTML
                $childToc = Spider-OneNote-Notebook -onenote $onenote -node $child -path $folder -notebookRoot $notebookRoot
                $tocHtml += $childToc
                $tocHtml += "</ul></li>`n"
            }

            # Store page level & name for next loop iteration
            $previouslevel = $child.pageLevel
            $previousname = $safeName 
        } catch {
            $errName = if ($child.name) { $child.name } else { "Unknown Node" }
            Log-Error "Failed while spidering node '$errName'. Error: $_"
        }
    }
    return $tocHtml
}

Function ReplaceIllegal {
    param ( $text )
    $illegal = [string]::join('',([System.IO.Path]::GetInvalidFileNameChars())) -replace '\\','\\'
    $replaced = $text -replace "[$illegal]",'_'
    return $replaced
}

Function Get-Folder($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null
    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select an export folder"
    $foldername.rootfolder = "MyComputer"
    if($foldername.ShowDialog() -eq "OK") { $folder += $foldername.SelectedPath }
    return $folder
}

# ================= MAIN EXECUTION =================

# Get export folder
$folder = Get-Folder

if (-not $folder) {
    Write-Host "No folder selected. Exiting."
    exit
}

# Define the global error log path at the root of the selected export directory
$script:errorLogPath = Join-Path -Path $folder -ChildPath "errors.log"

# Initialize the log file
Out-File -FilePath $script:errorLogPath -InputObject "=== OneNote Export Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===" -Encoding UTF8

try {
    # Connect to OneNote COM API
    $OneNote = New-Object -ComObject OneNote.Application
    [xml]$Hierarchy = ""
    $OneNote.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$Hierarchy)
} catch {
    Write-Host "CRITICAL ERROR: Failed to connect to OneNote COM API. Make sure OneNote Desktop is installed and open." -ForegroundColor Red
    Log-Error "Failed to initialize OneNote COM Object or GetHierarchy. Error: $_"
    exit
}

# Loop over each notebook
foreach ($notebook in $Hierarchy.Notebooks.Notebook ) {
    try {
        $name = ReplaceIllegal -text $notebook.name
        $nf = Join-Path -Path $folder -ChildPath $name
        Write-Host "=======================================" -ForegroundColor Magenta
        Write-Host "Notebook: $($nf)" -ForegroundColor Magenta
        Write-Host "=======================================" -ForegroundColor Magenta
        New-Item -Path $nf -ItemType directory -ErrorAction Ignore | Out-Null
        
        # Kick off the spidering and capture the generated HTML list
        $tocBody = Spider-OneNote-Notebook -onenote $OneNote -node $notebook -path $nf -notebookRoot $nf

        # Wrap the list in a clean, styled HTML document
        $safeNotebookName = [System.Net.WebUtility]::HtmlEncode($notebook.name)
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
        $nbName = if ($notebook.name) { $notebook.name } else { "Unknown" }
        Log-Error "Critical error processing notebook '$nbName'. Error: $_"
    }
}

# Cleanup filelist.xml files
Get-ChildItem -path $folder filelist.xml -Recurse | foreach { Remove-Item -Path $_.FullName -ErrorAction Ignore }

$finishTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Out-File -FilePath $script:errorLogPath -InputObject "=== OneNote Export Finished: $finishTime ===" -Append -Encoding UTF8

Write-Host "`nExport Complete! Check '$script:errorLogPath' for any errors that occurred." -ForegroundColor Cyan
