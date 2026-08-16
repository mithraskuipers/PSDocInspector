# DocInspector

A local, browser-based tool for keyword searching across office documents on Windows. Point it at a folder, give it keywords, and it scans `.docx`, `.doc`, `.xlsx`, `.xls`, `.pdf`, `.rtf`, and `.txt` files for matches — with OCR fallback for scanned PDFs.

Runs entirely on your machine as a small PowerShell HTTP server. No install, no external services, no data leaves your computer.

## Features

- Recursive folder scanning with live progress and streaming results
- Case-sensitive, whole-word, and regex search modes
- Legacy `.doc`/`.xls` support via Word/Excel automation (when installed), including Confluence-style HTML/MHTML `.doc` exports
- Flags files that can't be read and lets you re-run OCR on scanned PDFs from within the UI
- Filter and search within results by keyword, directory, or filename
- Export results to JSON or TXT, and re-import a previous JSON export
- One-click stop from the UI — no need to hunt down the terminal window

## Requirements

- Windows with PowerShell (Windows PowerShell or PowerShell 7+)
- Microsoft Word / Excel installed, only if you need to scan legacy `.doc` / `.xls` files
- An OCR engine available on the system, only if you need to OCR scanned PDFs

## Usage

Run `Run.bat`, or start it manually with an optional port:

```
Run.bat [port]
```

This launches the server (default port `8790`) and opens DocInspector in your default browser. From there:

1. Browse to or type a folder path.
2. Enter keywords (one per line, or comma-separated).
3. Set search options and click **Scan**.
4. Review results, filter them, expand rows for match previews, or export.

Click **Stop server** in the header to shut it down from the browser, or press `Ctrl+C` in the terminal window.

## Testing

- `run_smoke_tests.ps1` — quick standalone smoke tests, no dependencies:
  ```
  powershell -File run_smoke_tests.ps1
  ```
- `DocInspector_Tests.ps1` — Pester test suite:
  ```
  Invoke-Pester -Path .\DocInspector_Tests.ps1
  ```

Both dot-source `DocInspector.ps1` to test its extraction and matching functions directly, without starting the server.

## License

MIT — see [LICENSE](LICENSE).
