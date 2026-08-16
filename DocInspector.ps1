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

function Test-IsOleCompoundFile {
    <#
        Genuine binary .doc files are OLE Compound File Binary documents,
        which always start with the fixed 8-byte signature below. .doc files
        exported by systems like Confluence are actually HTML or MHTML text
        saved with a .doc extension, so they will never match this signature.
        Used to route .doc files to the correct extractor based on real
        content rather than trusting the extension. Fails safe (returns
        $false) on any read error so callers fall through to the HTML/MHTML
        path, which itself fails gracefully on genuinely unreadable content.
    #>
    param($Path)
    try {
        $sig = [byte[]]@(0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
        $buffer = New-Object byte[] 8
        $fs = [System.IO.File]::OpenRead($Path)
        try {
            $read = $fs.Read($buffer, 0, 8)
        } finally {
            $fs.Close()
        }
        if ($read -lt 8) { return $false }
        for ($i = 0; $i -lt 8; $i++) {
            if ($buffer[$i] -ne $sig[$i]) { return $false }
        }
        return $true
    } catch {
        return $false
    }
}

function ConvertFrom-QuotedPrintableBytes {
    <#
        Decodes a quoted-printable MIME body (as read via a 1-byte-per-char
        Latin-1/ISO-8859-1 string, so character offsets line up with raw file
        bytes) into the original byte array: "=XX" hex escapes become single
        bytes and soft line breaks ("=" followed by a line break) are removed.
        The caller re-decodes the resulting bytes using the part's declared
        charset, since quoted-printable text is charset-agnostic at this stage.
    #>
    param([string]$Text)
    $joined = $Text -replace '=\r?\n', ''
    $bytes = New-Object System.Collections.Generic.List[byte]
    $chars = $joined.ToCharArray()
    $i = 0
    while ($i -lt $chars.Length) {
        if ($chars[$i] -eq '=' -and ($i + 2) -lt $chars.Length -and
            $chars[($i + 1)] -match '[0-9A-Fa-f]' -and $chars[($i + 2)] -match '[0-9A-Fa-f]') {
            $hex = "$($chars[$i + 1])$($chars[$i + 2])"
            $bytes.Add([byte][Convert]::ToInt32($hex, 16))
            $i += 3
        } else {
            $bytes.Add([byte]$chars[$i])
            $i++
        }
    }
    return $bytes.ToArray()
}

function Get-EncodingByName {
    param([string]$Name)
    try {
        return [System.Text.Encoding]::GetEncoding($Name.Trim())
    } catch {
        return [System.Text.Encoding]::UTF8
    }
}

function ConvertTo-PlainTextFromHtml {
    <#
        Strips an HTML/Word-HTML fragment down to its human-readable text,
        so keyword matching and result previews only ever see what a user
        would actually read - never tags, attributes, CSS, JavaScript,
        Office/XML metadata, or MIME artifacts.
    #>
    param([string]$Html)
    if ([string]::IsNullOrEmpty($Html)) { return '' }

    $opts = [System.Text.RegularExpressions.RegexOptions]::Singleline
    $optsI = $opts -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase

    # Drop HTML comments first - this also removes Word's conditional
    # comment blocks (<!--[if gte mso 9]>...<![endif]-->) which commonly
    # wrap VML shapes and <xml> metadata that isn't visible document text.
    $t = [regex]::Replace($Html, '<!--.*?-->', ' ', $opts)

    # Drop entire elements whose content is never visible document text.
    $t = [regex]::Replace($t, '<script\b[^>]*>.*?</script>', ' ', $optsI)
    $t = [regex]::Replace($t, '<style\b[^>]*>.*?</style>', ' ', $optsI)
    $t = [regex]::Replace($t, '<xml\b[^>]*>.*?</xml>', ' ', $optsI)
    $t = [regex]::Replace($t, '<head\b[^>]*>.*?</head>', ' ', $optsI)

    # Turn block/line-break boundaries into newlines before stripping tags,
    # so words from separate elements don't run together.
    $t = [regex]::Replace($t, '(?i)<(br|/p|/div|/tr|/li|/h[1-6]|/table)\s*/?>', "`n")

    # Strip every remaining tag, including its attributes.
    $t = [regex]::Replace($t, '<[^>]+>', ' ')

    # Decode HTML entities (&amp;, &nbsp;, &#8217;, etc.)
    $t = [System.Net.WebUtility]::HtmlDecode($t)

    # Collapse whitespace produced by the tag stripping above.
    $t = [regex]::Replace($t, '[ \t]+', ' ')
    $t = [regex]::Replace($t, '\n[ \t]*', "`n")
    $t = [regex]::Replace($t, '\n{2,}', "`n")
    return $t.Trim()
}

function Get-DocHtmlText {
    <#
        Extracts human-readable text from a .doc file that is actually
        HTML or MHTML content (e.g. exported from Confluence) rather than a
        genuine binary Word document. Handles both a bare HTML/Word-HTML
        payload and a full MHTML multipart/related package, decoding
        quoted-printable/base64 bodies and the declared charset before
        stripping markup. Returns '' (not $null) on any parse failure so the
        caller reports "no readable text" instead of a misleading
        Word-unavailable message, and so a malformed file never crashes
        the scan.
    #>
    param($Path)
    try {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        if ($bytes.Length -eq 0) { return '' }
        $latin1 = [System.Text.Encoding]::GetEncoding('ISO-8859-1')
        $raw = $latin1.GetString($bytes)

        # Sanity check: if the content has none of the hallmarks of HTML or
        # MHTML, this .doc is neither a genuine binary Word document (already
        # ruled out by the OLE signature check) nor recognizable markup - it's
        # malformed or unsupported. Bail out cleanly rather than running markup
        # stripping over arbitrary bytes and reporting garbage as "text".
        if ($raw -notmatch '(?i)(<html|<body|<w:worddocument|multipart/related|mime-version:)') {
            return ''
        }

        $htmlFragment = $null

        if ($raw -match '(?im)^\s*(MIME-Version:|Content-Type:\s*multipart/related)') {
            $boundaryMatch = [regex]::Match($raw, '(?is)Content-Type:\s*multipart/related;.*?boundary\s*=\s*"?([^";\r\n]+)"?')
            if ($boundaryMatch.Success) {
                $boundary = $boundaryMatch.Groups[1].Value.Trim()
                if ($boundary) {
                    $splitToken = "--$boundary"
                    $parts = $raw -split [regex]::Escape($splitToken)
                    foreach ($part in $parts) {
                        if ($part -notmatch '(?im)^\s*Content-Type:\s*text/html') { continue }

                        $sep = $part.IndexOf("`r`n`r`n")
                        if ($sep -lt 0) { $sep = $part.IndexOf("`n`n") }
                        if ($sep -lt 0) { continue }

                        $headerBlock = $part.Substring(0, $sep)
                        $bodyBlock = $part.Substring($sep).TrimStart("`r", "`n")

                        $charset = 'UTF-8'
                        $csMatch = [regex]::Match($headerBlock, '(?i)charset\s*=\s*"?([\w\-]+)"?')
                        if ($csMatch.Success) { $charset = $csMatch.Groups[1].Value }

                        if ($headerBlock -match '(?im)^\s*Content-Transfer-Encoding:\s*quoted-printable') {
                            $bodyBytes = ConvertFrom-QuotedPrintableBytes -Text $bodyBlock
                        } elseif ($headerBlock -match '(?im)^\s*Content-Transfer-Encoding:\s*base64') {
                            try {
                                $bodyBytes = [Convert]::FromBase64String(($bodyBlock -replace '\s', ''))
                            } catch {
                                $bodyBytes = $latin1.GetBytes($bodyBlock)
                            }
                        } else {
                            $bodyBytes = $latin1.GetBytes($bodyBlock)
                        }

                        $enc = Get-EncodingByName -Name $charset
                        $htmlFragment = $enc.GetString($bodyBytes)
                        break
                    }
                }
            }
            if (-not $htmlFragment) {
                # Couldn't isolate the text/html MIME part (unexpected
                # boundary format, etc.) - fall back to the whole payload so
                # markup stripping still has a chance to recover the text
                # rather than reporting zero content outright.
                $htmlFragment = $raw
            }
        } else {
            # Not MHTML - treat as a bare HTML/Word-HTML document saved with
            # a .doc extension. Respect a declared <meta charset> if present.
            $charset = 'UTF-8'
            $csMatch = [regex]::Match($raw, '(?i)<meta[^>]+charset\s*=\s*"?([\w\-]+)')
            if ($csMatch.Success) { $charset = $csMatch.Groups[1].Value }
            $enc = Get-EncodingByName -Name $charset
            $htmlFragment = $enc.GetString($bytes)
        }

        return ConvertTo-PlainTextFromHtml -Html $htmlFragment
    } catch {
        return ''
    }
}

function Get-BinaryHeuristicText {
    <#
        Extracts human-readable text from legacy binary Office files (.doc,
        .xls) WITHOUT any dependency on Word/Excel being installed, licensed,
        or COM-automatable. Both formats store their text content as
        contiguous runs of either UTF-16LE code units or single-byte
        CP1252/ANSI characters, even though it's embedded inside a larger
        binary container (OLE Compound File structures, formatting records,
        etc). This walks the raw bytes and, at every position, greedily
        matches the longer of a UTF-16LE-printable run or an 8-bit-printable
        run, emits it as decoded text, and skips forward past non-text bytes.
        This is deliberately the same technique the classic 'strings' utility
        uses - it doesn't reconstruct document structure or reading order
        perfectly, but for keyword search purposes (the only thing this tool
        needs) it reliably recovers the actual words in the file regardless
        of whether Word is present, licensed, or currently able to start.

        IMPORTANT: the "printable" ranges below are deliberately narrow
        (essentially ASCII + Western-European Latin + common typographic
        punctuation), NOT "anything Unicode considers a letter". Random
        binary bytes (compressed streams, OLE structure tables, formatting
        records) reinterpreted as UTF-16LE code units land somewhere in the
        Unicode range purely by chance; if that range is allowed to include
        huge blocks like CJK Unified Ideographs or Hangul, binary noise gets
        misdecoded into what looks like real (but bogus) Chinese/Japanese/
        Korean text mixed into results. Keeping the allowed ranges tight to
        what Western business documents actually use avoids that entirely.
        Runs shorter than $MinRun are treated as noise (stray printable
        bytes inside binary formatting records) and discarded.
    #>
    param($Path, [int]$MinRun = 6)
    try {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
    } catch {
        return ''
    }
    $n = $bytes.Length
    if ($n -eq 0) { return '' }
    $ansi = [System.Text.Encoding]::GetEncoding(1252)

    # CP1252's 0x80-0x9F block is mostly either undefined or control
    # characters, with a handful of real typographic punctuation mixed in.
    # Allow only those specific known-good code points instead of the whole
    # block, so undefined/control bytes in that range aren't treated as text.
    $cp1252SpecialAllowed = @(0x80,0x82,0x83,0x84,0x85,0x86,0x87,0x88,0x89,0x8A,0x8B,0x8C,0x8E,0x91,0x92,0x93,0x94,0x95,0x96,0x97,0x98,0x99,0x9A,0x9B,0x9C,0x9E,0x9F)

    $sb = New-Object System.Text.StringBuilder
    $i = 0
    while ($i -lt $n) {
        # Longest run readable as UTF-16LE starting at $i. Allowed: tab/LF/CR,
        # ASCII printable, Latin-1 Supplement + Latin Extended-A/B (covers
        # accented Western/Central European text), and common typographic
        # punctuation (curly quotes, dashes, ellipsis, Euro/trademark signs).
        # Deliberately excludes CJK, Hangul, and other large scripts - see
        # comment above.
        $j = $i
        while (($j + 1) -lt $n) {
            $code = $bytes[$j] -bor ($bytes[$j + 1] -shl 8)
            $printable = ($code -eq 9 -or $code -eq 10 -or $code -eq 13) -or
                         ($code -ge 32 -and $code -le 126) -or
                         ($code -ge 0xA0 -and $code -le 0x24F) -or
                         ($code -ge 0x2010 -and $code -le 0x2027) -or
                         ($code -ge 0x2030 -and $code -le 0x2039) -or
                         ($code -eq 0x20AC -or $code -eq 0x2122)
            if (-not $printable) { break }
            $j += 2
        }
        $run16 = $j - $i

        # Longest run readable as 8-bit CP1252 starting at $i. Allowed:
        # tab/LF/CR, ASCII printable, the curated CP1252 0x80-0x9F subset
        # above, and the full 0xA0-0xFF Latin-1 Supplement block (accented
        # characters and common symbols, all genuinely printable in CP1252).
        $k = $i
        while ($k -lt $n) {
            $b = $bytes[$k]
            $printable8 = $b -eq 9 -or $b -eq 10 -or $b -eq 13 -or
                          ($b -ge 32 -and $b -le 126) -or
                          ($b -ge 0xA0) -or
                          ($cp1252SpecialAllowed -contains $b)
            if (-not $printable8) { break }
            $k++
        }
        $run8 = $k - $i

        if ($run16 -ge ($MinRun * 2) -and $run16 -ge $run8) {
            [void]$sb.Append([System.Text.Encoding]::Unicode.GetString($bytes, $i, $run16)).Append(' ')
            $i = $j
        } elseif ($run8 -ge $MinRun) {
            [void]$sb.Append($ansi.GetString($bytes, $i, $run8)).Append(' ')
            $i = $k
        } else {
            $i++
        }
    }
    return $sb.ToString()
}

function Get-DocText {
    <#
        Extracts text from a genuine binary .doc via the Word COM object.
        Returns $null on any failure. $script:LastDocOpenError distinguishes
        *why* for the caller: 'unavailable' means Word itself couldn't be
        launched (not installed/licensed), while any other value is the
        actual exception message from Documents.Open() - i.e. Word is fine,
        but this specific file could not be opened (locked, corrupted,
        Protected View, password-protected, etc). Get-FileText uses this to
        avoid reporting a misleading "Word is not available" reason when
        Word is actually working.
    #>
    param($Path)
    $script:LastDocOpenError = $null

    if ($script:WordFailed) {
        $script:LastDocOpenError = 'unavailable'
        return $null
    }
    if (-not $script:WordApp) {
        # Try twice: a fresh Click-to-Run session or a just-released COM
        # registration can throw a transient "Unable to cast COM object..."
        # / QueryInterface error on the very first activation attempt and
        # then succeed immediately after. One retry (with the object
        # released in between) resolves that without ever touching other
        # processes on the machine.
        $lastError = $null
        for ($attempt = 1; $attempt -le 2; $attempt++) {
            try {
                $script:WordApp = New-Object -ComObject Word.Application -Property @{ Visible = $false }
                $lastError = $null
                break
            } catch {
                $lastError = $_
                if ($script:WordApp) {
                    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($script:WordApp) | Out-Null } catch {}
                    $script:WordApp = $null
                }
                if ($attempt -lt 2) { Start-Sleep -Milliseconds 800 }
            }
        }
        if ($lastError) {
            $script:WordFailed = $true
            # Keep the actual COM error around separately from the
            # 'unavailable' sentinel so the caller can surface the real
            # reason (e.g. a registration/permission problem) instead of
            # always claiming Word isn't installed when it actually is.
            $script:WordUnavailableReason = $lastError.Exception.Message
            $script:LastDocOpenError = 'unavailable'
            return $null
        }
    }
    try {
        $doc = $script:WordApp.Documents.Open($Path, $false, $true, $false)
        $text = $doc.Content.Text
        $doc.Close($false)
        return $text
    } catch {
        $script:LastDocOpenError = $_.Exception.Message
        return $null
    }
}

function Get-XlsText {
    param($Path)
    if ($script:ExcelFailed) { return $null }
    if (-not $script:ExcelApp) {
        $lastError = $null
        for ($attempt = 1; $attempt -le 2; $attempt++) {
            try {
                $script:ExcelApp = New-Object -ComObject Excel.Application -Property @{ Visible = $false; DisplayAlerts = $false }
                $lastError = $null
                break
            } catch {
                $lastError = $_
                if ($script:ExcelApp) {
                    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($script:ExcelApp) | Out-Null } catch {}
                    $script:ExcelApp = $null
                }
                if ($attempt -lt 2) { Start-Sleep -Milliseconds 800 }
            }
        }
        if ($lastError) {
            $script:ExcelFailed = $true
            $script:ExcelUnavailableReason = $lastError.Exception.Message
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
        '.doc'  {
            # Not every .doc file is a binary Word document - systems like
            # Confluence export .doc files that are actually HTML or MHTML.
            # Detect the real format from content, not the extension.
            if (Test-IsOleCompoundFile -Path $Path) {
                # Primary path: byte-level heuristic extraction. Needs no
                # Word install/license/COM cooperation at all, so this alone
                # already covers the overwhelming majority of real .doc
                # files even on a machine where Word automation is fully
                # broken.
                $heuristicText = Get-BinaryHeuristicText -Path $Path

                # Word COM is now purely a bonus, tried only when the
                # heuristic came back thin (e.g. a file that's mostly
                # non-text runs it couldn't recover well), and only ever
                # used if it actually beats what the heuristic already
                # found. A broken/unavailable Word install can no longer
                # cause a file to be skipped - it can only fail to improve
                # on a result we already have.
                if ([string]::IsNullOrWhiteSpace($heuristicText) -or $heuristicText.Trim().Length -lt 40) {
                    $wordText = Get-DocText -Path $Path
                    if ($wordText -and $wordText.Trim().Length -gt $heuristicText.Trim().Length) {
                        return $wordText
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($heuristicText)) { return $heuristicText }

                # Heuristic found nothing usable either - try the HTML/MHTML
                # extractor as a last resort (covers OLE-wrapped files with
                # embedded HTML/XML content the byte scanner isn't tuned for).
                $htmlText = Get-DocHtmlText -Path $Path
                if (-not [string]::IsNullOrWhiteSpace($htmlText)) { return $htmlText }

                return $null
            } else {
                return Get-DocHtmlText -Path $Path
            }
        }
        '.xls'  {
            # Same heuristic-first approach as .doc: BIFF (.xls) binary
            # records also store cell strings as recoverable UTF-16LE/CP1252
            # runs, so this needs no Excel install at all for the common case.
            $heuristicText = Get-BinaryHeuristicText -Path $Path
            if ([string]::IsNullOrWhiteSpace($heuristicText) -or $heuristicText.Trim().Length -lt 40) {
                $excelText = Get-XlsText -Path $Path
                if ($excelText -and $excelText.Trim().Length -gt $heuristicText.Trim().Length) {
                    return $excelText
                }
            }
            return $heuristicText
        }
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

function Invoke-OpenFileRequest {
    <#
        Opens a scanned file in its default associated application (Word,
        Excel, Adobe Reader, etc.) via the OS, triggered by the user clicking
        a filename in the results table. Runs server-side because the
        browser itself has no way to launch a native app for a local path.
    #>
    param($Context)
    $req = (Get-RequestBody -Context $Context) | ConvertFrom-Json
    $fullPath = $req.fullPath

    if ([string]::IsNullOrWhiteSpace($fullPath) -or -not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        Send-Json -Context $Context -StatusCode 404 -Data @{ error = "File not found: $fullPath" }
        return
    }

    try {
        Start-Process -FilePath $fullPath | Out-Null
        Send-Json -Context $Context -Data @{ success = $true }
    } catch {
        Send-Json -Context $Context -StatusCode 500 -Data @{ error = "Could not open file: $($_.Exception.Message)" }
    }
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
    $sentResultsCount = 0
    $sentSkippedCount = 0
    $lastPartialSent = Get-Date

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
                    '.doc'  {
                        # Get-FileText already tried the HTML/MHTML fallback
                        # for .doc before giving up, so a $null here means a
                        # genuine binary .doc that Word could not produce
                        # text for. $script:LastDocOpenError (set by
                        # Get-DocText) tells us whether that's because Word
                        # itself is unavailable or because Word choked on
                        # this specific file - report the real reason
                        # instead of always blaming a missing Word install.
                        if ($script:LastDocOpenError -and $script:LastDocOpenError -ne 'unavailable') {
                            "Word could not open this file: $($script:LastDocOpenError)"
                        } elseif ($script:WordUnavailableReason) {
                            # Word is installed enough to be found on the
                            # system, but PowerShell couldn't instantiate it
                            # via COM - report the real error instead of a
                            # blanket "not installed" message.
                            "Microsoft Word could not be started - $($script:WordUnavailableReason)"
                        } else {
                            'Microsoft Word is not available on this machine - legacy .doc files require Word to be installed'
                        }
                    }
                    '.xls'  {
                        if ($script:ExcelUnavailableReason) {
                            "Microsoft Excel could not be started - $($script:ExcelUnavailableReason)"
                        } else {
                            'Microsoft Excel is not available on this machine - legacy .xls files require Excel to be installed'
                        }
                    }
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
                        # Wrapped in @() so a single match still serializes as
                        # a 1-element JSON array (ConvertTo-Json otherwise
                        # collapses a lone pipeline object to a bare scalar,
                        # which breaks .map()/.filter() on the client, most
                        # visibly when re-importing an exported JSON file).
                        snippets  = @($snippets | Select-Object -First 5)
                    }
                }
            }
        }

        $processed++
        Write-NdjsonLine -Stream $stream -Data @{
            type          = 'progress'
            processed     = $processed
            total         = $files.Count
            scanned       = $scanned
            skipped       = $skipped.Count
            currentFile   = $file.Name
            findingsCount = $results.Count
        }

        # Throttle full-row streaming to roughly every 2 seconds (rather than
        # every file) so a half-hour scan doesn't spend most of its time
        # serializing JSON - but still lets the person start reviewing real
        # findings and skipped-file reasons well before the scan finishes,
        # instead of staring at a bare progress bar the whole time. Only the
        # rows added since the last partial are sent; the client appends them.
        if (((Get-Date) - $lastPartialSent) -ge [TimeSpan]::FromSeconds(2) -or $processed -eq $files.Count) {
            if ($results.Count -gt $sentResultsCount -or $skipped.Count -gt $sentSkippedCount) {
                $newResults = if ($results.Count -gt $sentResultsCount) { @($results[$sentResultsCount..($results.Count - 1)]) } else { @() }
                $newSkipped = if ($skipped.Count -gt $sentSkippedCount) { @($skipped[$sentSkippedCount..($skipped.Count - 1)]) } else { @() }
                Write-NdjsonLine -Stream $stream -Data @{
                    type         = 'partial'
                    newResults   = $newResults
                    newSkipped   = $newSkipped
                    scannedFiles = $scanned
                    skippedCount = $skipped.Count
                    processed    = $processed
                }
                $sentResultsCount = $results.Count
                $sentSkippedCount = $skipped.Count
            }
            $lastPartialSent = Get-Date
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
    $sentResultsCount = 0
    $sentSkippedCount = 0
    $lastPartialSent = Get-Date

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
                        # Wrapped in @() so a single match still serializes as
                        # a 1-element JSON array (ConvertTo-Json otherwise
                        # collapses a lone pipeline object to a bare scalar,
                        # which breaks .map()/.filter() on the client, most
                        # visibly when re-importing an exported JSON file).
                        snippets  = @($snippets | Select-Object -First 5)
                        viaOcr    = $true
                    }
                }
            }
        }

        $processed++
        Write-NdjsonLine -Stream $stream -Data @{
            type          = 'progress'
            processed     = $processed
            total         = $items.Count
            scanned       = $scanned
            skipped       = $stillSkipped.Count
            currentFile   = $item.fileName
            findingsCount = $results.Count
        }

        if (((Get-Date) - $lastPartialSent) -ge [TimeSpan]::FromSeconds(2) -or $processed -eq $items.Count) {
            if ($results.Count -gt $sentResultsCount -or $stillSkipped.Count -gt $sentSkippedCount) {
                $newResults = if ($results.Count -gt $sentResultsCount) { @($results[$sentResultsCount..($results.Count - 1)]) } else { @() }
                $newSkipped = if ($stillSkipped.Count -gt $sentSkippedCount) { @($stillSkipped[$sentSkippedCount..($stillSkipped.Count - 1)]) } else { @() }
                Write-NdjsonLine -Stream $stream -Data @{
                    type         = 'partial'
                    newResults   = $newResults
                    newSkipped   = $newSkipped
                    scannedFiles = $scanned
                    skippedCount = $stillSkipped.Count
                    processed    = $processed
                }
                $sentResultsCount = $results.Count
                $sentSkippedCount = $stillSkipped.Count
            }
            $lastPartialSent = Get-Date
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
        '^/api/open-file$' {
            if ($method -eq 'POST') { Invoke-OpenFileRequest -Context $Context }
            else { $Context.Response.StatusCode = 405; $Context.Response.OutputStream.Close() }
        }
        default {
            $Context.Response.StatusCode = 404
            $Context.Response.OutputStream.Close()
        }
    }
}

# ---------- Main ----------

# Guard the listener startup so this script can be dot-sourced (e.g. by
# Pester tests exercising the extraction/matching functions above) without
# actually opening a port and launching a browser. Normal usage
# (".\DocInspector.ps1") is unaffected - $MyInvocation.InvocationName is only
# '.' when the script is dot-sourced.
if ($MyInvocation.InvocationName -eq '.') { return }

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Host ''
    Write-Host 'DocInspector must run in a Single-Threaded Apartment (STA) to automate Word/Excel.' -ForegroundColor Red
    Write-Host "This session is currently $([System.Threading.Thread]::CurrentThread.GetApartmentState()), which will make every legacy .doc/.xls file fail." -ForegroundColor Red
    Write-Host 'Relaunch with:' -ForegroundColor Yellow
    Write-Host "  pwsh -sta -File `"$PSCommandPath`""
    Write-Host "  (or) powershell.exe -sta -File `"$PSCommandPath`""
    Write-Host ''
    exit 1
}

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
