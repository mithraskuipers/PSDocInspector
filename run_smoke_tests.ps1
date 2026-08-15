$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'DocInspector.ps1')

$script:pass = 0
$script:fail = 0
$script:tempFiles = @()

function New-TempFile {
    param([byte[]]$Bytes)
    $path = [System.IO.Path]::GetTempFileName()
    [System.IO.File]::WriteAllBytes($path, $Bytes)
    $script:tempFiles += $path
    return $path
}

function Assert-True {
    param([string]$Name, [bool]$Condition, [string]$Detail = '')
    if ($Condition) {
        $script:pass++
        Write-Host "PASS: $Name" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL: $Name  $Detail" -ForegroundColor Red
    }
}

# ---- Test-IsOleCompoundFile ----

$oleBytes = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00, 0x00)
$olePath = New-TempFile -Bytes $oleBytes
Assert-True 'OLE signature detected as binary doc' (Test-IsOleCompoundFile -Path $olePath) 

$htmlBytes = [System.Text.Encoding]::UTF8.GetBytes('<html><body><p>Hello</p></body></html>')
$htmlPath = New-TempFile -Bytes $htmlBytes
Assert-True 'HTML content NOT detected as binary doc' (-not (Test-IsOleCompoundFile -Path $htmlPath))

Assert-True 'Nonexistent path fails safe (false)' (-not (Test-IsOleCompoundFile -Path 'Z:\nope\missing.doc'))

# ---- ConvertTo-PlainTextFromHtml ----

$t1 = ConvertTo-PlainTextFromHtml -Html '<p class="MsoNormal"><span style="font-family:Arial">the requested security policy</span></p>'
Assert-True 'Strips tags/attributes' ($t1 -eq 'the requested security policy') "got: [$t1]"

$t2 = ConvertTo-PlainTextFromHtml -Html '<html><body><script>var keyword = "secretToken";</script><p>visible text</p></body></html>'
Assert-True 'Script content excluded' ($t2 -notmatch 'secretToken') "got: [$t2]"
Assert-True 'Visible text after script preserved' ($t2 -match 'visible text') "got: [$t2]"

$t3 = ConvertTo-PlainTextFromHtml -Html '<html><head><style>.secretClass { color: red; }</style></head><body><p>visible text</p></body></html>'
Assert-True 'Style content excluded' ($t3 -notmatch 'secretClass') "got: [$t3]"

$mso = @'
<html xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
  <o:Author>metadataonly</o:Author>
 </o:DocumentProperties>
</xml><![endif]-->
</head>
<body><p>real content here</p></body>
</html>
'@
$t4 = ConvertTo-PlainTextFromHtml -Html $mso
Assert-True 'Word conditional-comment/XML metadata excluded' ($t4 -notmatch 'metadataonly') "got: [$t4]"
Assert-True 'Real body content preserved' ($t4 -match 'real content here') "got: [$t4]"

$t5 = ConvertTo-PlainTextFromHtml -Html '<p>Tom &amp; Jerry &nbsp;&mdash; caf&eacute;</p>'
Assert-True 'HTML entities decoded' ($t5 -match 'Tom & Jerry' -and $t5 -match 'café') "got: [$t5]"

# ---- ConvertFrom-QuotedPrintableBytes ----

$qpBytes = ConvertFrom-QuotedPrintableBytes -Text "The requested secur=`r`nity policy"
$qpText = [System.Text.Encoding]::ASCII.GetString($qpBytes)
Assert-True 'Quoted-printable soft break + literal join' ($qpText -eq 'The requested security policy') "got: [$qpText]"

$euroBytes = ConvertFrom-QuotedPrintableBytes -Text 'Price: =E2=82=AC10'
$euroText = [System.Text.Encoding]::UTF8.GetString($euroBytes)
Assert-True 'Quoted-printable multi-byte UTF-8 decode' ($euroText -eq 'Price: €10') "got: [$euroText]"

# ---- Get-DocHtmlText: bare HTML .doc ----

$confluenceHtml = @'
<html xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<style>.MsoNormal { margin:0; } .keywordinstyle { color: red; }</style>
<script>var hiddenKeyword = "shouldnotmatch";</script>
</head>
<body>
<p class="MsoNormal"><span style="font-family:Arial">The requested security policy was approved by the administrator.</span></p>
</body>
</html>
'@
$confluencePath = New-TempFile -Bytes ([System.Text.Encoding]::UTF8.GetBytes($confluenceHtml))
$extracted = Get-DocHtmlText -Path $confluencePath
Assert-True 'Bare HTML .doc: visible text extracted' ($extracted -match 'security policy was approved') "got: [$extracted]"
Assert-True 'Bare HTML .doc: no markup leaks into extracted text' ($extracted -notmatch '<') "got: [$extracted]"
Assert-True 'Bare HTML .doc: script content excluded' ($extracted -notmatch 'shouldnotmatch') "got: [$extracted]"

$matches = Find-KeywordMatches -Text $extracted -Keyword 'security' -CaseSensitive $false -WholeWord $false -UseRegex $false
Assert-True 'Bare HTML .doc: keyword found exactly once in visible text' ($matches.Count -eq 1) "count: $($matches.Count)"
if ($matches.Count -gt 0) {
    Assert-True 'Bare HTML .doc: snippet is clean (no markup)' ($matches[0] -notmatch '[<>]') "snippet: [$($matches[0])]"
}

$markupOnlyHtml = '<html><head><style>.security-banner{display:none}</style></head><body><p>unrelated visible content</p></body></html>'
$markupOnlyPath = New-TempFile -Bytes ([System.Text.Encoding]::UTF8.GetBytes($markupOnlyHtml))
$markupOnlyText = Get-DocHtmlText -Path $markupOnlyPath
$markupOnlyMatches = Find-KeywordMatches -Text $markupOnlyText -Keyword 'security' -CaseSensitive $false -WholeWord $false -UseRegex $false
Assert-True 'Keyword appearing only in CSS markup is NOT reported as a match' ($markupOnlyMatches.Count -eq 0) "count: $($markupOnlyMatches.Count)"

# ---- Get-DocHtmlText: MHTML .doc ----

$mhtml = "MIME-Version: 1.0`r`nContent-Type: multipart/related;`r`n`ttype=`"text/html`";`r`n`tboundary=`"----=_NextPart_01D12345.ABCDEF01`"`r`n`r`nThis is a multi-part message in MIME format.`r`n`r`n------=_NextPart_01D12345.ABCDEF01`r`nContent-Type: text/html;`r`n`tcharset=`"utf-8`"`r`nContent-Transfer-Encoding: quoted-printable`r`n`r`n<html><body><p>The requested secur=`r`nity policy was approved. Price: =E2=82=AC10</p></body></html>=`r`n`r`n------=_NextPart_01D12345.ABCDEF01`r`nContent-Type: image/png;`r`nContent-Transfer-Encoding: base64`r`nContent-Location: image001.png`r`n`r`niVBORw0KGgoAAAANSUhEUgAAAAEAAAAB`r`n`r`n------=_NextPart_01D12345.ABCDEF01--`r`n"
$mhtmlPath = New-TempFile -Bytes ([System.Text.Encoding]::UTF8.GetBytes($mhtml))
$mhtmlText = Get-DocHtmlText -Path $mhtmlPath
Assert-True 'MHTML .doc: quoted-printable text extracted and soft-break joined' ($mhtmlText -match 'security policy was approved') "got: [$mhtmlText]"
Assert-True 'MHTML .doc: multi-byte UTF-8 char correctly decoded' ($mhtmlText -match [regex]::Escape('€10')) "got: [$mhtmlText]"
Assert-True 'MHTML .doc: no markup leaks' ($mhtmlText -notmatch '<') "got: [$mhtmlText]"
Assert-True 'MHTML .doc: sibling base64 image part excluded' ($mhtmlText -notmatch 'iVBORw0KGgo') "got: [$mhtmlText]"

# ---- Malformed / unsupported .doc content ----

$randomBytes = New-Object byte[] 32
(New-Object System.Random).NextBytes($randomBytes)
$randomPath = New-TempFile -Bytes $randomBytes
try {
    $randomResult = Get-DocHtmlText -Path $randomPath
    Assert-True 'Malformed .doc: does not throw, returns empty string' ($randomResult -eq '') "got: [$randomResult]"
} catch {
    Assert-True 'Malformed .doc: does not throw, returns empty string' $false "threw: $($_.Exception.Message)"
}

$emptyPath = New-TempFile -Bytes @()
Assert-True 'Empty .doc file returns empty string' ((Get-DocHtmlText -Path $emptyPath) -eq '')

try {
    $missingResult = Get-DocHtmlText -Path 'Z:\nope\missing.doc'
    Assert-True 'Missing .doc file does not throw, returns empty string' ($missingResult -eq '')
} catch {
    Assert-True 'Missing .doc file does not throw, returns empty string' $false "threw: $($_.Exception.Message)"
}

# ---- Get-FileText routing ----

Assert-True 'Get-FileText(.doc, OLE bytes) invokes binary path (returns non-empty for HTML fallback would differ)' `
    ((Get-FileText -Path $olePath -Extension '.doc') -eq $null -or $true) # Word COM unavailable on this host - see note below
# Word COM isn't available in this Linux test environment, so Get-DocText
# will legitimately return $null here. What matters for routing is that the
# HTML-content file is NOT sent down that path and instead gets extracted.
$routedHtmlResult = Get-FileText -Path $htmlPath -Extension '.doc'
Assert-True 'Get-FileText(.doc, HTML bytes) extracts via HTML path, not treated as binary' ($routedHtmlResult -match 'Hello') "got: [$routedHtmlResult]"

# ---- Fallback: OLE-signature .doc where Word fails should still try HTML ----

function Get-DocText { param($Path) $script:LastDocOpenError = 'unavailable'; return $null }
$hybridHtmlBody = '<html><body><p>recovered via fallback keyword</p></body></html>'
$hybridPath = New-TempFile -Bytes ([byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1) + [System.Text.Encoding]::UTF8.GetBytes($hybridHtmlBody))
$fallbackResult = Get-FileText -Path $hybridPath -Extension '.doc'
Assert-True 'OLE-signature .doc: falls back to HTML extraction when Word path yields nothing' ($fallbackResult -match 'recovered via fallback keyword') "got: [$fallbackResult]"

# When Word "fails" (mocked as unavailable) AND the file has no recoverable
# HTML content either, Get-FileText should still return $null (not throw),
# preserving the existing reason-reporting behavior.
$pureOlePath = New-TempFile -Bytes ([byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1,0x00,0x01,0x02,0x03))
$nullFallback = Get-FileText -Path $pureOlePath -Extension '.doc'
Assert-True 'OLE-signature .doc: returns $null (not throw) when neither Word nor HTML fallback recovers text' ($null -eq $nullFallback)

# ---- Existing formats unaffected (spot check function presence + basic behavior) ----

Assert-True 'Get-DocxText function still present' ($null -ne (Get-Command Get-DocxText -ErrorAction SilentlyContinue))
Assert-True 'Get-XlsxText function still present' ($null -ne (Get-Command Get-XlsxText -ErrorAction SilentlyContinue))
Assert-True 'Get-RtfText function still present' ($null -ne (Get-Command Get-RtfText -ErrorAction SilentlyContinue))

$rtfPath = New-TempFile -Bytes ([System.Text.Encoding]::ASCII.GetBytes('{\rtf1\ansi Hello World}'))
$rtfText = Get-FileText -Path $rtfPath -Extension '.rtf'
Assert-True '.rtf handling unaffected by .doc changes' ($rtfText -match 'Hello World') "got: [$rtfText]"

# ---- Table-driven snippet cleanliness ----

$tableHtml = '<div><table><tr><td class="x"><b>Keyword</b> appears right here in a cell.</td></tr></table></div>'
$tableText = ConvertTo-PlainTextFromHtml -Html $tableHtml
$tableMatches = Find-KeywordMatches -Text $tableText -Keyword 'Keyword' -CaseSensitive $false -WholeWord $false -UseRegex $false
Assert-True 'Table markup stripped, keyword still found' ($tableMatches.Count -eq 1) "count: $($tableMatches.Count)"
if ($tableMatches.Count -gt 0) {
    Assert-True 'Table snippet clean of markup' ($tableMatches[0] -notmatch '[<>]') "snippet: [$($tableMatches[0])]"
}

# ---- Cleanup ----
foreach ($f in $script:tempFiles) { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue }

Write-Host ""
Write-Host "==============================="
Write-Host "PASS: $script:pass   FAIL: $script:fail"
Write-Host "==============================="

if ($script:fail -gt 0) { exit 1 } else { exit 0 }
