param(
    [int]$Port = 8790
)

$ErrorActionPreference = 'Stop'
$WwwRoot = $PSScriptRoot
$script:WordApp = $null
$script:ExcelApp = $null
$script:OcrUnavailableReason = $null

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

function Expand-FlateBytes {
    param([byte[]]$Bytes)
    if (-not $Bytes -or $Bytes.Length -lt 3) { return $null }

    # PDF FlateDecode streams are zlib-wrapped (RFC 1950): a 2-byte header
    # followed by raw DEFLATE data and a 4-byte Adler-32 trailer. .NET's
    # DeflateStream only understands the raw DEFLATE payload (RFC 1951), so
    # the zlib header is stripped before decompressing. If that fails (e.g.
    # a non-standard header), fall back to treating the bytes as raw DEFLATE
    # with no header at all rather than throwing.
    foreach ($offset in @(2, 0)) {
        if ($offset -ge $Bytes.Length) { continue }
        $input = $null
        $deflate = $null
        $output = $null
        try {
            $input = New-Object System.IO.MemoryStream(, $Bytes)
            [void]$input.Seek($offset, [System.IO.SeekOrigin]::Begin)
            $deflate = New-Object System.IO.Compression.DeflateStream($input, [System.IO.Compression.CompressionMode]::Decompress)
            $output = New-Object System.IO.MemoryStream
            $deflate.CopyTo($output)
            if ($output.Length -gt 0) { return $output.ToArray() }
        } catch {
            # Corrupt/unsupported stream data - try the next offset (or give
            # up and let the caller treat this stream as unavailable).
        } finally {
            if ($deflate) { $deflate.Close() }
            if ($input) { $input.Close() }
        }
    }
    return $null
}

function Get-PdfText {
    param($Path)
    try {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        $encoding = [System.Text.Encoding]::GetEncoding('ISO-8859-1')
        $raw = $encoding.GetString($bytes)

        # Start from the raw file text so existing behavior for uncompressed
        # content streams (literal "(text) Tj" operators sitting directly in
        # the file) is preserved unchanged.
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append($raw)

        # Locate stream objects whose dictionary declares /FlateDecode,
        # decompress their data with DeflateStream, and append the
        # decompressed text so the same Tj/TJ extraction below also picks up
        # text from compressed content streams (the common case for PDFs
        # produced by Word, Chrome, Adobe, etc).
        $streamPattern = New-Object System.Text.RegularExpressions.Regex(
            '<<(?<dict>(?:[^<>]|<<[^<>]*>>)*)>>\s*stream\r?\n(?<data>.*?)\r?\n?endstream',
            [System.Text.RegularExpressions.RegexOptions]::Singleline
        )

        foreach ($m in $streamPattern.Matches($raw)) {
            $dict = $m.Groups['dict'].Value
            if ($dict -notmatch '/FlateDecode') { continue }

            try {
                $streamBytes = $encoding.GetBytes($m.Groups['data'].Value)
                $decompressed = Expand-FlateBytes -Bytes $streamBytes
                if ($decompressed) {
                    [void]$sb.Append(' ').Append($encoding.GetString($decompressed))
                }
            } catch {
                # Decompression failed for this stream (corrupt/unsupported) -
                # skip it and keep processing the rest of the document
                # instead of failing the whole extraction.
                continue
            }
        }

        $combined = $sb.ToString()
        $out = New-Object System.Text.StringBuilder
        foreach ($m in [regex]::Matches($combined, '\(((?:[^()\\]|\\.)*)\)\s*Tj')) {
            $s = $m.Groups[1].Value -replace '\\\(', '(' -replace '\\\)', ')' -replace '\\\\', '\'
            [void]$out.Append($s).Append(' ')
        }
        foreach ($m in [regex]::Matches($combined, '\[((?:[^\[\]])*)\]\s*TJ')) {
            foreach ($im in [regex]::Matches($m.Groups[1].Value, '\(((?:[^()\\]|\\.)*)\)')) {
                $s = $im.Groups[1].Value -replace '\\\(', '(' -replace '\\\)', ')' -replace '\\\\', '\'
                [void]$out.Append($s)
            }
            [void]$out.Append(' ')
        }
        return $out.ToString()
    } catch { return '' }
}

# ---------- OCR (Windows.Media.Ocr + Windows.Data.Pdf - built into Windows 10+) ----------

# Lazily resolves the generic AsTask<T>(IAsyncOperation<T>) and AsTask(IAsyncAction)
# extension methods so WinRT async calls can be awaited synchronously from PowerShell,
# which has no native "await". Resolved once and cached at script scope.
function Initialize-WinRtAwait {
    if ($script:AsTaskWithResult -and $script:AsTaskNoResult) { return }
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $methods = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 }
    $script:AsTaskWithResult = ($methods | Where-Object { $_.IsGenericMethod -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
    $script:AsTaskNoResult   = ($methods | Where-Object { -not $_.IsGenericMethod -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
}

function Wait-WinRtOperation {
    param($Operation, [Type]$ResultType)
    Initialize-WinRtAwait
    $method = $script:AsTaskWithResult.MakeGenericMethod($ResultType)
    $task = $method.Invoke($null, @($Operation))
    $task.Wait() | Out-Null
    return $task.Result
}

function Wait-WinRtAction {
    param($Action)
    Initialize-WinRtAwait
    $task = $script:AsTaskNoResult.Invoke($null, @($Action))
    $task.Wait() | Out-Null
}

# Checks (once) whether the Windows OCR + PDF rendering WinRT APIs are present
# and an OCR engine can be created for one of the user's installed languages.
# Result is cached in script scope, mirroring the Word/Excel COM availability checks.
function Test-OcrAvailable {
    if ($script:OcrChecked) { return $script:OcrReady }
    $script:OcrChecked = $true
    $script:OcrReady = $false

    # The WinRT type projections this feature relies on (Windows.Media.Ocr,
    # Windows.Data.Pdf, etc.) are only available under Windows PowerShell 5.1.
    # PowerShell 7/Core dropped WinRT support, so fail fast here with a clear,
    # user-facing reason instead of letting a cryptic type-resolution error
    # surface later.
    if ($PSVersionTable.PSEdition -eq 'Core') {
        $script:OcrUnavailableReason = 'OCR requires Windows PowerShell 5.1 (powershell.exe). This server was started with PowerShell 7/Core, which does not support the Windows Runtime APIs OCR depends on - restart it with powershell.exe instead of pwsh.exe.'
        Write-Host "OCR unavailable: $($script:OcrUnavailableReason)" -ForegroundColor Yellow
        return $false
    }

    try {
        [void][Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime]
        [void][Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
        [void][Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
        [void][Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
        [void][Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]

        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
        if (-not $engine) {
            $script:OcrUnavailableReason = 'No OCR language pack is installed for any of your Windows display languages. Install one via Settings > Time & language > Language & region > (your language) > Options > Optical character recognition.'
            Write-Host "OCR unavailable: $($script:OcrUnavailableReason)" -ForegroundColor Yellow
            return $false
        }
        $script:OcrEngine = $engine
        $script:OcrReady = $true
    } catch {
        $script:OcrUnavailableReason = "OCR unavailable: $($_.Exception.Message)"
        Write-Host $script:OcrUnavailableReason -ForegroundColor Yellow
        $script:OcrReady = $false
    }
    return $script:OcrReady
}

# Rasterizes each page of a PDF (via Windows.Data.Pdf) and runs Windows' built-in
# OCR engine (Windows.Media.Ocr) against the rendered image, concatenating the
# recognized text. Used only as a fallback for PDFs with no extractable text layer
# (scanned/image-only PDFs). Returns $null if OCR itself is unavailable, or throws
# on a genuine per-file failure so the caller can record a reason.
function Get-PdfTextViaOcr {
    param($Path)
    if (-not (Test-OcrAvailable)) { return $null }

    $storageFile = Wait-WinRtOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) ([Windows.Storage.StorageFile])
    $pdfDoc = Wait-WinRtOperation ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($storageFile)) ([Windows.Data.Pdf.PdfDocument])

    $sb = New-Object System.Text.StringBuilder
    $scale = 2.0 # upscale rendering for better OCR accuracy on small text

    for ($i = 0; $i -lt $pdfDoc.PageCount; $i++) {
        $page = $pdfDoc.GetPage([uint32]$i)
        try {
            $ras = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
            $renderOptions = New-Object Windows.Data.Pdf.PdfPageRenderOptions
            $renderOptions.DestinationWidth  = [uint32]([Math]::Ceiling($page.Size.Width  * $scale))
            $renderOptions.DestinationHeight = [uint32]([Math]::Ceiling($page.Size.Height * $scale))
            Wait-WinRtAction ($page.RenderToStreamAsync($ras, $renderOptions))

            $decoder = Wait-WinRtOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($ras)) ([Windows.Graphics.Imaging.BitmapDecoder])
            $softwareBitmap = Wait-WinRtOperation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
            $ocrResult = Wait-WinRtOperation ($script:OcrEngine.RecognizeAsync($softwareBitmap)) ([Windows.Media.Ocr.OcrResult])

            if ($ocrResult -and $ocrResult.Text) {
                [void]$sb.Append($ocrResult.Text).Append(' ')
            }
        } finally {
            $page.Dispose()
        }
    }

    return $sb.ToString()
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
            $ocrEligible = ($ext -eq '.pdf') -and ($reason -eq 'No extractable text found - likely a scanned or image-based PDF')
            $skipped += [PSCustomObject]@{
                fileName     = $file.Name
                directory    = $file.DirectoryName
                fullPath     = $file.FullName
                extension    = $ext
                reason       = $reason
                ocrEligible  = $ocrEligible
                ocrAttempted = $false
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

# Re-runs keyword matching against a specific set of previously-skipped PDFs
# using OCR text extraction (Get-PdfTextViaOcr) instead of the normal content-
# stream parsing. Only intended for PDFs the client flagged as ocrEligible
# (i.e. "no extractable text - likely scanned/image-based"), not encrypted or
# genuinely corrupted files.
function Invoke-OcrRequest {
    param($Context)
    $req = (Get-RequestBody -Context $Context) | ConvertFrom-Json

    $items = @($req.items)
    $keywords = @($req.keywords) | Where-Object { $_ -and $_.Trim() -ne '' } | ForEach-Object { $_.Trim() } | Select-Object -Unique
    $caseSensitive = [bool]$req.caseSensitive
    $wholeWord = [bool]$req.wholeWord
    $useRegex = [bool]$req.useRegex

    if ($items.Count -eq 0) {
        Send-Json -Context $Context -StatusCode 400 -Data @{ error = 'No skipped PDFs were provided for OCR' }
        return
    }
    if ($keywords.Count -eq 0) {
        Send-Json -Context $Context -StatusCode 400 -Data @{ error = 'No keywords provided' }
        return
    }
    if (-not (Test-OcrAvailable)) {
        $reason = if ($script:OcrUnavailableReason) { $script:OcrUnavailableReason } else { 'OCR is not available on this machine.' }
        Send-Json -Context $Context -StatusCode 400 -Data @{ error = $reason }
        return
    }

    $Context.Response.ContentType = 'application/x-ndjson; charset=utf-8'
    $Context.Response.SendChunked = $true
    $stream = $Context.Response.OutputStream

    Write-NdjsonLine -Stream $stream -Data @{ type = 'start'; totalFiles = $items.Count }

    $results = @()
    $stillSkipped = @()
    $recoveredPaths = @()
    $scanned = 0
    $processed = 0

    foreach ($item in $items) {
        $fullPath = $item.fullPath
        $reason = $null
        $text = $null

        if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
            $reason = 'File no longer exists at the recorded path'
        } else {
            try {
                $text = Get-PdfTextViaOcr -Path $fullPath
            } catch {
                $reason = "OCR error: $($_.Exception.Message)"
            }
        }

        if (-not $reason -and [string]::IsNullOrWhiteSpace($text)) {
            $reason = 'OCR completed but found no extractable text - the page(s) may be blank or unreadable'
        }

        if ($reason) {
            $stillSkipped += [PSCustomObject]@{
                fileName     = $item.fileName
                directory    = $item.directory
                fullPath     = $fullPath
                extension    = $item.extension
                reason       = $reason
                ocrEligible  = $false
                ocrAttempted = $true
            }
        } else {
            $scanned++
            $recoveredPaths += $fullPath
            foreach ($kw in $keywords) {
                $snippets = Find-KeywordMatches -Text $text -Keyword $kw -CaseSensitive $caseSensitive -WholeWord $wholeWord -UseRegex $useRegex
                if ($snippets.Count -gt 0) {
                    $results += [PSCustomObject]@{
                        keyword   = $kw
                        fileName  = $item.fileName
                        directory = $item.directory
                        fullPath  = $fullPath
                        extension = $item.extension
                        count     = $snippets.Count
                        snippets  = $snippets | Select-Object -First 5
                        viaOcr    = $true
                    }
                }
            }
        }

        $processed++
        Write-NdjsonLine -Stream $stream -Data @{
            type        = 'progress'
            processed   = $processed
            total       = $items.Count
            scanned     = $scanned
            currentFile = $item.fileName
        }
    }

    Write-NdjsonLine -Stream $stream -Data @{
        type           = 'done'
        results        = $results
        skippedFiles   = $stillSkipped
        recoveredPaths = $recoveredPaths
        scannedFiles   = $scanned
        totalFiles     = $items.Count
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
        '^/api/ocr-status$' {
            $ready = Test-OcrAvailable
            Send-Json -Context $Context -Data @{ available = $ready; reason = $(if (-not $ready) { $script:OcrUnavailableReason }) }
        }
        '^/api/scan$' {
            if ($method -eq 'POST') { Invoke-ScanRequest -Context $Context }
            else { $Context.Response.StatusCode = 405; $Context.Response.OutputStream.Close() }
        }
        '^/api/ocr$' {
            if ($method -eq 'POST') { Invoke-OcrRequest -Context $Context }
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
