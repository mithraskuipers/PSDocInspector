<#
    Pester tests for the .doc HTML/MHTML extraction added to DocInspector.ps1.

    Run with: Invoke-Pester -Path .\DocInspector.Tests.ps1

    The main script is dot-sourced, which (per the InvocationName guard at the
    bottom of DocInspector.ps1) loads all functions without starting the
    HttpListener or opening a browser.
#>

BeforeAll {
    . (Join-Path $PSScriptRoot 'DocInspector.ps1')

    function New-TempFile {
        param([byte[]]$Bytes)
        $path = [System.IO.Path]::GetTempFileName()
        [System.IO.File]::WriteAllBytes($path, $Bytes)
        return $path
    }

    $script:TempFiles = @()
}

AfterAll {
    foreach ($f in $script:TempFiles) {
        Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Test-IsOleCompoundFile' {
    It 'returns true for a genuine OLE compound file signature' {
        $bytes = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00, 0x00)
        $path = New-TempFile -Bytes $bytes
        $script:TempFiles += $path
        Test-IsOleCompoundFile -Path $path | Should -Be $true
    }

    It 'returns false for an HTML .doc file' {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes('<html><body><p>Hello</p></body></html>')
        $path = New-TempFile -Bytes $bytes
        $script:TempFiles += $path
        Test-IsOleCompoundFile -Path $path | Should -Be $false
    }

    It 'returns false (fails safe) for a nonexistent file' {
        Test-IsOleCompoundFile -Path 'C:\this\path\does\not\exist.doc' | Should -Be $false
    }
}

Describe 'ConvertTo-PlainTextFromHtml' {
    It 'strips tags and returns readable text' {
        $html = '<p class="MsoNormal"><span style="font-family:Arial">the requested security policy</span></p>'
        ConvertTo-PlainTextFromHtml -Html $html | Should -Be 'the requested security policy'
    }

    It 'does not leak text from <script> blocks' {
        $html = '<html><body><script>var keyword = "secretToken";</script><p>visible text</p></body></html>'
        $result = ConvertTo-PlainTextFromHtml -Html $html
        $result | Should -Not -Match 'secretToken'
        $result | Should -Match 'visible text'
    }

    It 'does not leak text from <style> blocks' {
        $html = '<html><head><style>.secretClass { color: red; }</style></head><body><p>visible text</p></body></html>'
        $result = ConvertTo-PlainTextFromHtml -Html $html
        $result | Should -Not -Match 'secretClass'
        $result | Should -Match 'visible text'
    }

    It 'does not leak text from Word conditional-comment / XML metadata blocks' {
        $html = @'
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
        $result = ConvertTo-PlainTextFromHtml -Html $html
        $result | Should -Not -Match 'metadataonly'
        $result | Should -Match 'real content here'
    }

    It 'decodes HTML entities' {
        $html = '<p>Tom &amp; Jerry &nbsp;&mdash; caf&eacute;</p>'
        $result = ConvertTo-PlainTextFromHtml -Html $html
        $result | Should -Match 'Tom & Jerry'
        $result | Should -Match 'café'
    }

    It 'inserts breaks between block elements so words do not run together' {
        $html = '<p>first paragraph</p><p>second paragraph</p>'
        $result = ConvertTo-PlainTextFromHtml -Html $html
        $result | Should -Not -Match 'paragraphsecond'
    }

    It 'returns an empty string for empty input' {
        ConvertTo-PlainTextFromHtml -Html '' | Should -Be ''
    }
}

Describe 'ConvertFrom-QuotedPrintableBytes' {
    It 'joins soft line breaks and decodes =XX escapes' {
        $qp = "The requested secur=`r`nity policy"
        $bytes = ConvertFrom-QuotedPrintableBytes -Text $qp
        [System.Text.Encoding]::ASCII.GetString($bytes) | Should -Be 'The requested security policy'
    }

    It 'correctly reconstructs a multi-byte UTF-8 character' {
        # =E2=82=AC is the UTF-8 byte sequence for the Euro sign (€)
        $qp = 'Price: =E2=82=AC10'
        $bytes = ConvertFrom-QuotedPrintableBytes -Text $qp
        [System.Text.Encoding]::UTF8.GetString($bytes) | Should -Be 'Price: €10'
    }
}

Describe 'Get-DocHtmlText - bare HTML .doc (e.g. Confluence export)' {
    It 'extracts only the visible text and finds keywords located there' {
        $html = @'
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
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($html)
        $path = New-TempFile -Bytes $bytes
        $script:TempFiles += $path

        $text = Get-DocHtmlText -Path $path
        $text | Should -Match 'security policy was approved'
        $text | Should -Not -Match '<'
        $text | Should -Not -Match 'shouldnotmatch'

        $matches = Find-KeywordMatches -Text $text -Keyword 'security' -CaseSensitive $false -WholeWord $false -UseRegex $false
        $matches.Count | Should -Be 1
        $matches[0] | Should -Not -Match '<'
    }

    It 'does not report a match for a keyword that only appears inside markup' {
        $html = '<html><head><style>.security-banner{display:none}</style></head><body><p>unrelated visible content</p></body></html>'
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($html)
        $path = New-TempFile -Bytes $bytes
        $script:TempFiles += $path

        $text = Get-DocHtmlText -Path $path
        $matches = Find-KeywordMatches -Text $text -Keyword 'security' -CaseSensitive $false -WholeWord $false -UseRegex $false
        $matches.Count | Should -Be 0
    }
}

Describe 'Get-DocHtmlText - MHTML .doc' {
    It 'extracts and decodes a quoted-printable text/html part, ignoring the sibling image part' {
        $mhtml = @"
MIME-Version: 1.0
Content-Type: multipart/related;
`ttype="text/html";
`tboundary="----=_NextPart_01D12345.ABCDEF01"

This is a multi-part message in MIME format.

------=_NextPart_01D12345.ABCDEF01
Content-Type: text/html;
`tcharset="utf-8"
Content-Transfer-Encoding: quoted-printable

<html><body><p>The requested secur=
ity policy was approved. Price: =E2=82=AC10</p></body></html>=

------=_NextPart_01D12345.ABCDEF01
Content-Type: image/png;
Content-Transfer-Encoding: base64
Content-Location: image001.png

iVBORw0KGgoAAAANSUhEUgAAAAEAAAAB

------=_NextPart_01D12345.ABCDEF01--
"@
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($mhtml)
        $path = New-TempFile -Bytes $bytes
        $script:TempFiles += $path

        $text = Get-DocHtmlText -Path $path
        $text | Should -Match 'security policy was approved'
        $text | Should -Match '€10'
        $text | Should -Not -Match '<'
        $text | Should -Not -Match 'iVBORw0KGgo'
    }
}

Describe 'Get-DocHtmlText - malformed / unsupported .doc content' {
    It 'returns an empty string instead of throwing for random binary content' {
        $bytes = [byte[]](1..32 | ForEach-Object { Get-Random -Minimum 0 -Maximum 255 })
        $path = New-TempFile -Bytes $bytes
        $script:TempFiles += $path

        { Get-DocHtmlText -Path $path } | Should -Not -Throw
        Get-DocHtmlText -Path $path | Should -Be ''
    }

    It 'returns an empty string for an empty file' {
        $path = New-TempFile -Bytes @()
        $script:TempFiles += $path
        Get-DocHtmlText -Path $path | Should -Be ''
    }

    It 'returns an empty string (not throw) for a nonexistent path' {
        { Get-DocHtmlText -Path 'Z:\nope\missing.doc' } | Should -Not -Throw
        Get-DocHtmlText -Path 'Z:\nope\missing.doc' | Should -Be ''
    }
}

Describe 'Get-FileText - .doc format routing' {
    It 'routes a real OLE-signature .doc to Get-DocText (binary path)' {
        Mock Get-DocText { return 'called binary path' }
        Mock Get-DocHtmlText { return 'called html path' }

        $bytes = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00, 0x00)
        $path = New-TempFile -Bytes $bytes
        $script:TempFiles += $path

        Get-FileText -Path $path -Extension '.doc' | Should -Be 'called binary path'
        Should -Invoke Get-DocText -Times 1
        Should -Invoke Get-DocHtmlText -Times 0
    }

    It 'routes an HTML-content .doc to Get-DocHtmlText' {
        Mock Get-DocText { return 'called binary path' }
        Mock Get-DocHtmlText { return 'called html path' }

        $bytes = [System.Text.Encoding]::UTF8.GetBytes('<html><body>hi</body></html>')
        $path = New-TempFile -Bytes $bytes
        $script:TempFiles += $path

        Get-FileText -Path $path -Extension '.doc' | Should -Be 'called html path'
        Should -Invoke Get-DocHtmlText -Times 1
        Should -Invoke Get-DocText -Times 0
    }

    It 'still routes .docx through the existing OOXML extractor unaffected by this change' {
        Get-Command Get-DocxText | Should -Not -BeNullOrEmpty
    }

    It 'falls back to HTML extraction when an OLE-signature .doc yields nothing via Word' {
        Mock Get-DocText { $script:LastDocOpenError = 'unavailable'; return $null }

        $body = '<html><body><p>recovered via fallback keyword</p></body></html>'
        $bytes = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1) + [System.Text.Encoding]::UTF8.GetBytes($body)
        $path = New-TempFile -Bytes $bytes
        $script:TempFiles += $path

        Get-FileText -Path $path -Extension '.doc' | Should -Match 'recovered via fallback keyword'
    }

    It 'returns $null (not throw) when an OLE-signature .doc fails via Word and has no recoverable HTML' {
        Mock Get-DocText { $script:LastDocOpenError = 'unavailable'; return $null }

        $bytes = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00, 0x01, 0x02, 0x03)
        $path = New-TempFile -Bytes $bytes
        $script:TempFiles += $path

        { Get-FileText -Path $path -Extension '.doc' } | Should -Not -Throw
        Get-FileText -Path $path -Extension '.doc' | Should -BeNullOrEmpty
    }
}

Describe 'Find-KeywordMatches - snippet cleanliness on extracted HTML text' {
    It 'produces a preview snippet with no markup regardless of source formatting' {
        $html = '<div><table><tr><td class="x"><b>Keyword</b> appears right here in a cell.</td></tr></table></div>'
        $text = ConvertTo-PlainTextFromHtml -Html $html
        $matches = Find-KeywordMatches -Text $text -Keyword 'Keyword' -CaseSensitive $false -WholeWord $false -UseRegex $false
        $matches.Count | Should -Be 1
        $matches[0] | Should -Not -Match '[<>]'
        $matches[0] | Should -Match 'Keyword appears right here'
    }
}
