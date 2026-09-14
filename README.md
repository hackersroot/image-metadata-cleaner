# File Metadata Cleaner

A Python desktop GUI for batch metadata cleanup. Select a folder, include subfolders if needed, then choose **Clean and export**. Original file formats and relative folder structure are preserved.

## Install and launch

Requires Python 3.10+ with Tkinter (included in standard Windows Python installations).

```powershell
python -m pip install -r requirements.txt
python image_metadata_cleaner.py
```

On Windows, after installing dependencies, you can also double-click `launch_cleaner.bat`. The existing Python filename is retained for compatibility; the window is now named File Metadata Cleaner.

## Export options

- **Cleaned folder**: select a parent directory. The tool creates a fresh `cleaned_files_...` folder. If the destination is blank, it uses the source folder. Originals remain intact.
- **ZIP archive**: choose a new `.zip` filename. Includes successfully cleaned files and a CSV report. Existing ZIP files are never overwritten.
- **Replace originals**: overwrite only files that clean and verify successfully, after the GUI confirmation. Export destination is ignored. Keep your own backup.
- **Export CSV report**: save results after any run, including replace mode. Records each filename, status and error/detail. Folder/ZIP exports include the report automatically.

Stopping finishes the current file, then exports completed results and records remaining files as unprocessed. Reports include filenames; handle them accordingly when sharing. Exports preserve file types; there is no document or image format conversion.

## Supported formats and precise cleaning scope

| Formats | Removed | Preserved / limits |
| --- | --- | --- |
| JPG/JPEG, PNG, WebP, BMP, TIFF | EXIF/GPS, XMP, comments, PNG text, embedded thumbnails, ICC profiles | Static images only. Applies orientation first. JPEG re-encoded at quality 95; ICC removal may change appearance. Required image-structure fields remain. |
| PDF | Document Info properties, document ID, reachable XMP metadata streams; discarded metadata objects are omitted from rewritten output | Pages, text, forms and other document content retained. Encrypted PDFs and signature fields are rejected. Annotation authors/comments, attachments and metadata inside embedded files/images are not scrubbed. |
| DOCX, XLSX, PPTX | Standard `docProps` XML properties, including author/title/dates/company/custom properties; property thumbnails; ZIP comments/extra fields and source entry timestamps | Document content parts are kept byte-for-byte. Comments, revisions, revision IDs, custom XML, embedded objects and media metadata are not scrubbed. Signed packages are rejected. Encrypted/non-ZIP packages fail safely. |
| Legacy DOC, XLS, PPT | Standard OLE SummaryInformation and DocumentSummaryInformation property streams, including their custom-property sections | Body streams remain unchanged. Other embedded properties, revision history, directory timestamps and unused disk sectors are not scrubbed. This is limited property cleanup, not complete anonymization. |

Unsupported files, including GIF, HEIC, RAW, RTF, ODT and macro-enabled Office extensions, are reported as skipped. Animated/multipage supported image types are reported as failures with an explanation rather than losing frames. Malformed inputs are reported individually. Previous `cleaned_files_...` and `cleaned_images_...` directories and symbolic links are excluded during recursive scans.

“Removed” means properties are absent/empty, not the literal text `null`. Filesystem names/timestamps, visible personal information, sidecars and application caches are outside this tool's scope. Office applications may regenerate properties when a file is saved again. Legacy signature/encryption detection is limited; use unsigned, unencrypted legacy inputs. Property removal can affect fields that reference those properties.

Each file is written to a temporary file, reopened to verify targeted metadata removal, and atomically moved into place. Failed files do not replace originals. A filesystem or export failure is shown in the status. Very large files can consume significant memory; OOXML packages over 1 GB uncompressed are rejected.

## Verification performed

The included regression suite exercises:

- All five image encoders, EXIF orientation, PNG transparency, animation rejection and damaged input safety.
- PDF Info/XMP removal, absence of seeded metadata in output bytes, page text, exact before/after rendered pixels and an interactive form field.
- Encrypted and signature-bearing PDF rejection without replacing the source.
- DOCX/XLSX/PPTX reopening, body-part byte equality, content/formula preservation, standard/custom property removal and signed-package rejection.
- Legacy property-stream removal using synthetic OLE fixtures and unchanged body streams. These are not full Microsoft Office application compatibility tests.
- GUI construction, the actual Start/worker/event flow, folder/ZIP/replace exports, cancellation, CSV output and recursive output exclusion.

Run the tests:

```powershell
python -m pip install -r requirements-test.txt
python test_metadata_cleaner.py
```

The tests create temporary fixtures and clean them up. They do not process your personal files.
