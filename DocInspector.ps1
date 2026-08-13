param(
    [int]$Port = 8790
)

$ErrorActionPreference = 'Stop'
$WwwRoot = $PSScriptRoot
$script:WordApp = $null
$script:ExcelApp = $null

# ---------- Text extraction ----------

function Get-DocxText {
    param($Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry('word/document.xml')
        if (-not $entry) { return '' }
        $reader = New-Object System.IO.StreamReader($entry.Open())
        $xml = $reader.ReadToEnd()
        $reader.Close()
        $text = [regex]::Replace($xml, '<[^>]+>', ' ')
        return [System.Net.WebUtility]::HtmlDecode($text)
    } finally { $zip.Dispose() }
}

function Get-XlsxText {
    param($Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $sb = New-Object System.Text.StringBuilder
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName -match '^xl/(worksheets/sheet\d+\.xml|sharedStrings\.xml)$') {
                $reader = New-Object System.IO.StreamReader($entry.Open())
                $xml = $reader.ReadToEnd()
                $reader.Close()
                $text = [regex]::Replace($xml, '<[^>]+>', ' ')
                [void]$sb.AppendLine([System.Net.WebUtility]::HtmlDecode($text))
            }
        }
        return $sb.ToString()
    } finally { $zip.Dispose() }
}

function Get-DocText {
    param($Path)
    if ($script:WordFailed) { return $null }
    if (-not $script:WordApp) {
        try {
            $script:WordApp = New-Object -ComObject Word.Application
            $script:WordApp.Visible = $false
        } catch {
            $script:WordFailed = $true
            return $null
        }
    }
    try {
        $doc = $script:WordApp.Documents.Open($Path, $false, $true, $false)
        $text = $doc.Content.Text
        $doc.Close($false)
        return $text
    } catch { return $null }
}

function Get-XlsText {
    param($Path)
    if ($script:ExcelFailed) { return $null }
    if (-not $script:ExcelApp) {
        try {
            $script:ExcelApp = New-Object -ComObject Excel.Application
            $script:ExcelApp.Visible = $false
            $script:ExcelApp.DisplayAlerts = $false
        } catch {
            $script:ExcelFailed = $true
            return $null
        }
    }
    try {
        $wb = $script:ExcelApp.Workbooks.Open($Path, $false, $true)
        $sb = New-Object System.Text.StringBuilder
        foreach ($ws in $wb.Worksheets) {
            $vals = $ws.UsedRange.Value2
            if ($vals) {
                foreach ($v in $vals) { if ($null -ne $v) { [void]$sb.Append($v).Append(' ') } }
            }
        }
        $wb.Close($false)
        return $sb.ToString()
    } catch { return $null }
}

function Get-PdfText {
    param($Path)
    try {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        $raw = [System.Text.Encoding]::GetEncoding('ISO-8859-1').GetString($bytes)
        $sb = New-Object System.Text.StringBuilder
        foreach ($m in [regex]::Matches($raw, '\(((?:[^()\\]|\\.)*)\)\s*Tj')) {
            $s = $m.Groups[1].Value -replace '\\\(', '(' -replace '\\\)', ')' -replace '\\\\', '\'
            [void]$sb.Append($s).Append(' ')
        }
        foreach ($m in [regex]::Matches($raw, '\[((?:[^\[\]])*)\]\s*TJ')) {
            foreach ($im in [regex]::Matches($m.Groups[1].Value, '\(((?:[^()\\]|\\.)*)\)')) {
                $s = $im.Groups[1].Value -replace '\\\(', '(' -replace '\\\)', ')' -replace '\\\\', '\'
                [void]$sb.Append($s)
            }
            [void]$sb.Append(' ')
        }
        return $sb.ToString()
    } catch { return '' }
}

function Get-RtfText {
    param($Path)
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $text = [regex]::Replace($raw, '\\[a-zA-Z]+-?\d* ?', ' ')
        $text = $text -replace '[{}]', ''
        return ($text -replace '\\\r?\n', ' ')
    } catch { return '' }
}

function Get-FileText {
    param($Path, $Extension)
    switch ($Extension) {
        '.docx' { return Get-DocxText -Path $Path }
        '.xlsx' { return Get-XlsxText -Path $Path }
        '.doc'  { return Get-DocText -Path $Path }
        '.xls'  { return Get-XlsText -Path $Path }
        '.pdf'  { return Get-PdfText -Path $Path }
        '.rtf'  { return Get-RtfText -Path $Path }
        '.txt'  { return Get-Content -LiteralPath $Path -Raw -ErrorAction SilentlyContinue }
        default { return '' }
    }
}

# ---------- Keyword matching ----------

function Find-KeywordMatches {
    param($Text, $Keyword, [bool]$CaseSensitive, [bool]$WholeWord, [bool]$UseRegex)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }

    $options = [System.Text.RegularExpressions.RegexOptions]::None
    if (-not $CaseSensitive) { $options = $options -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }

    if ($UseRegex) {
        $pattern = $Keyword
    } else {
        $escaped = [regex]::Escape($Keyword)
        $pattern = if ($WholeWord) { "\b$escaped\b" } else { $escaped }
    }

    try {
        $regex = New-Object System.Text.RegularExpressions.Regex($pattern, $options)
    } catch { return @() }

    $results = @()
    foreach ($m in $regex.Matches($Text)) {
        $start = [Math]::Max(0, $m.Index - 40)
        $len = [Math]::Min(80 + $m.Length, $Text.Length - $start)
        $snippet = ($Text.Substring($start, $len).Trim() -replace '\s+', ' ')
        $results += $snippet
    }
    return $results
}

# ---------- HTTP helpers ----------

function Get-RequestBody {
    param($Context)
    $reader = New-Object System.IO.StreamReader($Context.Request.InputStream, [System.Text.Encoding]::UTF8)
    $body = $reader.ReadToEnd()
    $reader.Close()
    return $body
}

function Send-Json {
    param($Context, $Data, [int]$StatusCode = 200)
    $json = $Data | ConvertTo-Json -Depth 10 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $Context.Response.StatusCode = $StatusCode
    $Context.Response.ContentType = 'application/json; charset=utf-8'
    $Context.Response.ContentLength64 = $bytes.Length
    $Context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Context.Response.OutputStream.Close()
}

function Write-NdjsonLine {
    param($Stream, $Data)
    $json = $Data | ConvertTo-Json -Depth 10 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json + "`n")
    $Stream.Write($bytes, 0, $bytes.Length)
    $Stream.Flush()
}

function Send-StaticFile {
    param($Context, $Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        Write-Host "File not found: $Path" -ForegroundColor Yellow
        $Context.Response.StatusCode = 404
        $bytes = [System.Text.Encoding]::UTF8.GetBytes("Not found on server: $Path`nMake sure index.html and app.js are in the same folder as DocInspector.ps1.")
        $Context.Response.ContentType = 'text/plain; charset=utf-8'
        $Context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $Context.Response.OutputStream.Close()
        return
    }
    $ext = [System.IO.Path]::GetExtension($Path).ToLower()
    $contentType = switch ($ext) {
        '.html' { 'text/html; charset=utf-8' }
        '.js'   { 'application/javascript; charset=utf-8' }
        '.css'  { 'text/css; charset=utf-8' }
        default { 'application/octet-stream' }
    }
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $Context.Response.ContentType = $contentType
    $Context.Response.ContentLength64 = $bytes.Length
    $Context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Context.Response.OutputStream.Close()
}

function Show-FolderBrowserDialog {
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $rs.Open()

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        if (-not ([System.Management.Automation.PSTypeName]'Native.ForceForeground').Type) {
            Add-Type -Namespace Native -Name ForceForeground -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
[DllImport("user32.dll")] public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
[DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
'@
        }

        $owner = New-Object System.Windows.Forms.Form
        $owner.TopMost = $true
        $owner.ShowInTaskbar = $false
        $owner.StartPosition = 'CenterScreen'
        $owner.Size = New-Object System.Drawing.Size(1, 1)
        $owner.Opacity = 0
        $owner.Show()
        $owner.Activate()

        # Windows blocks background processes from stealing foreground focus.
        # Attach our input thread to the current foreground window's thread so
        # SetForegroundWindow is allowed to actually work, then detach after.
        $fgWindow = [Native.ForceForeground]::GetForegroundWindow()
        $fgThreadId = 0
        [void][Native.ForceForeground]::GetWindowThreadProcessId($fgWindow, [ref]$fgThreadId)
        $curThreadId = [Native.ForceForeground]::GetCurrentThreadId()
        $attached = $false
        if ($fgThreadId -ne 0 -and $fgThreadId -ne $curThreadId) {
            $attached = [Native.ForceForeground]::AttachThreadInput($curThreadId, $fgThreadId, $true)
        }
        [Native.ForceForeground]::ShowWindow($owner.Handle, 5) | Out-Null   # SW_SHOW
        [Native.ForceForeground]::SetForegroundWindow($owner.Handle) | Out-Null
        if ($attached) {
            [void][Native.ForceForeground]::AttachThreadInput($curThreadId, $fgThreadId, $false)
        }

        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = 'Select folder to scan'
        $dialog.ShowNewFolderButton = $false
        $result = $dialog.ShowDialog($owner)
        $owner.Close()
        $owner.Dispose()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $dialog.SelectedPath
        }
    })

    try {
        $output = $ps.Invoke()
        if ($ps.HadErrors) {
            foreach ($e in $ps.Streams.Error) { Write-Host "Folder dialog error: $e" -ForegroundColor Yellow }
        }
        if ($output -and $output.Count -gt 0) { return [string]$output[0] }
        return $null
    } catch {
        Write-Host "Folder dialog error: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    } finally {
        $ps.Dispose()
        $rs.Close()
        $rs.Dispose()
    }
}

# ---------- Scan ----------

function Invoke-ScanRequest {
    param($Context)
    $req = (Get-RequestBody -Context $Context) | ConvertFrom-Json

    $rootPath = $req.path
    $recursive = [bool]$req.recursive
    $keywords = @($req.keywords) | Where-Object { $_ -and $_.Trim() -ne '' } | ForEach-Object { $_.Trim() } | Select-Object -Unique
    $caseSensitive = [bool]$req.caseSensitive
    $wholeWord = [bool]$req.wholeWord
    $useRegex = [bool]$req.useRegex

    if (-not (Test-Path -LiteralPath $rootPath -PathType Container)) {
        Send-Json -Context $Context -StatusCode 400 -Data @{ error = "Folder not found: $rootPath" }
        return
    }
    if ($keywords.Count -eq 0) {
        Send-Json -Context $Context -StatusCode 400 -Data @{ error = 'No keywords provided' }
        return
    }

    $supportedExt = @('.docx', '.doc', '.xlsx', '.xls', '.pdf', '.rtf', '.txt')
    $gciParams = @{ LiteralPath = $rootPath; File = $true; ErrorAction = 'SilentlyContinue' }
    if ($recursive) { $gciParams['Recurse'] = $true }

    $files = Get-ChildItem @gciParams | Where-Object {
        $ext = $_.Extension.ToLower()
        ($supportedExt -contains $ext) -and ($ext -ne '.md') -and ($_.BaseName -notmatch '^readme')
    }

    # Stream progress to the client as newline-delimited JSON (NDJSON) so the
    # UI can show a live progress bar while the scan (which can take a while
    # for large folders / Office COM automation) is still running.
    $Context.Response.ContentType = 'application/x-ndjson; charset=utf-8'
    $Context.Response.SendChunked = $true
    $stream = $Context.Response.OutputStream

    Write-NdjsonLine -Stream $stream -Data @{ type = 'start'; totalFiles = $files.Count }

    $results = @()
    $skipped = @()
    $scanned = 0
    $processed = 0

    foreach ($file in $files) {
        $ext = $file.Extension.ToLower()
        $text = $null
        $reason = $null

        try {
            $text = Get-FileText -Path $file.FullName -Extension $ext
        } catch {
            $reason = "Error reading file: $($_.Exception.Message)"
        }

        if (-not $reason) {
            if ($null -eq $text) {
                $reason = switch ($ext) {
                    '.doc'  { 'Microsoft Word is not available on this machine - legacy .doc files require Word to be installed' }
                    '.xls'  { 'Microsoft Excel is not available on this machine - legacy .xls files require Excel to be installed' }
                    '.pdf'  { 'Could not extract text - the PDF may be encrypted, corrupted, or use an unsupported structure' }
                    '.rtf'  { 'Could not read RTF content' }
                    default { 'Could not extract text from this file' }
                }
            } elseif ([string]::IsNullOrWhiteSpace($text)) {
                $reason = if ($ext -eq '.pdf') {
                    'No extractable text found - likely a scanned or image-based PDF'
                } else {
                    'File appears to contain no readable text'
                }
            }
        }

        if ($reason) {
            $skipped += [PSCustomObject]@{
                fileName  = $file.Name
                directory = $file.DirectoryName
                fullPath  = $file.FullName
                extension = $ext
                reason    = $reason
            }
        } else {
            $scanned++
            foreach ($kw in $keywords) {
                $snippets = Find-KeywordMatches -Text $text -Keyword $kw -CaseSensitive $caseSensitive -WholeWord $wholeWord -UseRegex $useRegex
                if ($snippets.Count -gt 0) {
                    $results += [PSCustomObject]@{
                        keyword   = $kw
                        fileName  = $file.Name
                        directory = $file.DirectoryName
                        fullPath  = $file.FullName
                        extension = $ext
                        count     = $snippets.Count
                        snippets  = $snippets | Select-Object -First 5
                    }
                }
            }
        }

        $processed++
        Write-NdjsonLine -Stream $stream -Data @{
            type        = 'progress'
            processed   = $processed
            total       = $files.Count
            scanned     = $scanned
            currentFile = $file.Name
        }
    }

    Write-NdjsonLine -Stream $stream -Data @{
        type          = 'done'
        results       = $results
        skippedFiles  = $skipped
        scannedFiles  = $scanned
        totalFiles    = $files.Count
        sourcePath    = $rootPath
    }
    $stream.Close()
}

# ---------- Routing ----------

function Handle-Request {
    param($Context)
    $path = $Context.Request.Url.AbsolutePath
    $method = $Context.Request.HttpMethod

    switch -Regex ($path) {
        '^/$'           { Send-StaticFile -Context $Context -Path (Join-Path $WwwRoot 'index.html') }
        '^/app\.js$'    { Send-StaticFile -Context $Context -Path (Join-Path $WwwRoot 'app.js') }
        '^/api/browse$' {
            $selected = Show-FolderBrowserDialog
            Send-Json -Context $Context -Data @{ path = $selected }
        }
        '^/api/scan$' {
            if ($method -eq 'POST') { Invoke-ScanRequest -Context $Context }
            else { $Context.Response.StatusCode = 405; $Context.Response.OutputStream.Close() }
        }
        default {
            $Context.Response.StatusCode = 404
            $Context.Response.OutputStream.Close()
        }
    }
}

# ---------- Main ----------

$listener = New-Object System.Net.HttpListener
$prefix = "http://localhost:$Port/"
$listener.Prefixes.Add($prefix)
$listener.Start()
Write-Host "DocInspector running at $prefix"
Start-Process $prefix

try {
    while ($listener.IsListening) {
        try {
            $context = $listener.GetContext()
        } catch {
            Write-Host "GetContext error (ignored, still listening): $($_.Exception.Message)" -ForegroundColor Yellow
            continue
        }
        try {
            Handle-Request -Context $context
        } catch {
            Write-Host "Request error: $($_.Exception.Message)" -ForegroundColor Yellow
            try {
                $context.Response.StatusCode = 500
                $bytes = [System.Text.Encoding]::UTF8.GetBytes("Server error: $($_.Exception.Message)")
                $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
                $context.Response.OutputStream.Close()
            } catch {}
        }
    }
} finally {
    if ($script:WordApp) {
        $script:WordApp.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($script:WordApp) | Out-Null
    }
    if ($script:ExcelApp) {
        $script:ExcelApp.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($script:ExcelApp) | Out-Null
    }
    $listener.Stop()
}
