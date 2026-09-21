# Publisher to PDF Converter

A simple one-click tool to batch-convert Microsoft Publisher (`.pub`) files to PDF on Windows, using Publisher's own automation — no separate converter or online upload needed.

## What it does

- Scans a folder (and all its subfolders) for `.pub` files and converts each to a PDF saved right next to it
- Skips any file that already has a PDF, so it's safe to run again later
- Opens each file read-only, in its own fresh Publisher process, so one broken file can't derail the rest of the batch
- Keeps a dated log, a short report, and (if anything is left unconverted) a to-do list, all saved in a `PublisherToPDF-logs` folder next to the script
- Automatically falls back to an alternative automation method if Publisher's usual interface isn't available on a particular PC

## Requirements

- Windows, with Microsoft Publisher installed
- PowerShell (already included with Windows)

## How to use it

1. Download `PublishertoPDF.bat` and place it in the folder you want to convert (it will also check every subfolder).
2. Close Microsoft Publisher if it's open.
3. Double-click `PublishertoPDF.bat`.
4. If a Publisher "Security Notice" pop-up appears, click **Disable** (not Enable).
5. When it finishes, check the `PublisherToPDF-logs` folder for the report.

### Converting only specific files

To convert just a few files instead of scanning the whole folder, put their full paths in a plain text file (one per line) and drag that file onto `PublishertoPDF.bat`.

## License

Released under the MIT License — see [LICENSE](LICENSE) for details.
