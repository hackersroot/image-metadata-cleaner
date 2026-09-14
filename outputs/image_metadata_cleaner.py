"""GUI batch image and document property cleaner. See requirements.txt."""
from pathlib import Path
import os
import csv
import shutil
import struct
import zipfile
import xml.etree.ElementTree as ET
import queue
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageOps

EXTENSIONS = {'.jpg', '.jpeg', '.png', '.webp', '.bmp', '.tif', '.tiff'}
FORMATS = {'JPEG', 'PNG', 'WEBP', 'BMP', 'TIFF'}


def clean_image(source: Path, destination: Path):
    """Rebuild a static image from pixels; atomically write a verified clean file.

    Does not copy EXIF, GPS, XMP, comments, text, thumbnails, or ICC profiles.
    Encoder-required structural fields are allowed. Animated/multipage files
    are rejected so that frames/pages are never silently lost.
    """
    with Image.open(source) as original:
        fmt = original.format
        if fmt not in FORMATS:
            raise ValueError(f'Unsupported image format: {fmt}')
        if getattr(original, 'n_frames', 1) != 1:
            raise ValueError('Animated or multipage image skipped (preserves frames/pages)')
        original.load()
        oriented = ImageOps.exif_transpose(original)
        # Expand palettes and transparency keys into pixels before dropping info.
        if oriented.mode == 'P' or 'transparency' in oriented.info:
            oriented = oriented.convert('RGBA')
        if fmt == 'JPEG' and oriented.mode not in ('RGB', 'L', 'CMYK'):
            oriented = oriented.convert('RGB')
        if fmt == 'WEBP':
            oriented = oriented.convert('RGBA' if 'A' in oriented.getbands() else 'RGB')
        clean = Image.frombytes(oriented.mode, oriented.size, oriented.tobytes())
        options = {}
        if fmt == 'JPEG':
            options = {'quality': 95, 'subsampling': 0}
        elif fmt == 'WEBP':
            options = {'lossless': True}
        elif fmt == 'TIFF':
            options = {'compression': 'tiff_deflate'}
        destination.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(prefix='.metadata-clean-', dir=destination.parent)
        os.close(fd)
        try:
            clean.save(temporary, format=fmt, **options)
            with Image.open(temporary) as check:
                check.load()
                # TIFF stores required image structure in the EXIF-like IFD.
                if fmt != 'TIFF' and check.getexif():
                    raise ValueError('Verification failed: EXIF remains')
                forbidden = {'exif', 'xmp', 'XML:com.adobe.xmp', 'icc_profile', 'comment'}
                if forbidden.intersection(check.info):
                    raise ValueError('Verification failed: metadata remains')
                if fmt == 'PNG' and getattr(check, 'text', {}):
                    raise ValueError('Verification failed: PNG text remains')
            os.replace(temporary, destination)
        finally:
            clean.close()
            if os.path.exists(temporary):
                os.unlink(temporary)


DOCUMENTS = {'.pdf', '.docx', '.xlsx', '.pptx', '.doc', '.xls', '.ppt'}
ALL_EXTENSIONS = EXTENSIONS | DOCUMENTS


def clean_pdf(source, destination):
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import DictionaryObject, ArrayObject, IndirectObject, NameObject
    reader = PdfReader(source)
    if reader.is_encrypted:
        raise ValueError('Encrypted PDF is unsupported; decrypt a copy first')
    writer = PdfWriter(clone_from=reader)
    seen = set()

    def scrub(obj):
        if isinstance(obj, IndirectObject):
            obj = obj.get_object()
        if id(obj) in seen:
            return
        seen.add(id(obj))
        if isinstance(obj, DictionaryObject):
            if obj.get('/FT') == '/Sig' or obj.get('/Type') == '/Sig':
                raise ValueError('Signed PDF skipped to protect its signature')
            obj.pop(NameObject('/Metadata'), None)
            for child in list(obj.values()):
                scrub(child)
        elif isinstance(obj, ArrayObject):
            for child in obj:
                scrub(child)
    scrub(writer.root_object)
    writer.metadata = None
    writer._ID = None
    # A removed reference alone can leave the old XMP serialized as an orphan.
    # Retain only objects reachable from the cleaned catalog (pypdf 6.x).
    for index, obj in enumerate(writer._objects):
        if obj is not None and id(obj) not in seen:
            writer._objects[index] = None
    writer.write(destination)
    check = PdfReader(destination)
    if check.metadata or check.trailer.get('/ID'):
        raise ValueError('PDF property verification failed')
    seen.clear()
    def verify(obj):
        if isinstance(obj, IndirectObject):
            obj = obj.get_object()
        if id(obj) in seen:
            return
        seen.add(id(obj))
        if isinstance(obj, DictionaryObject):
            if '/Metadata' in obj:
                raise ValueError('PDF XMP verification failed')
            for child in obj.values():
                verify(child)
        elif isinstance(obj, ArrayObject):
            for child in obj:
                verify(child)
    verify(check.trailer['/Root'])
    if len(check.pages) != len(reader.pages):
        raise ValueError('PDF page count changed')


def clean_office(source, destination):
    # Preserve document/content XML bytes exactly; blank property parts only.
    with zipfile.ZipFile(source) as zin:
        names = zin.namelist()
        if len(names) != len(set(names)):
            raise ValueError('Ambiguous ZIP with duplicate members')
        if any(n.lower().startswith('_xmlsignatures/') for n in names):
            raise ValueError('Signed Office document skipped')
        expected = {'.docx': 'word/document.xml', '.xlsx': 'xl/workbook.xml', '.pptx': 'ppt/presentation.xml'}
        if expected[source.suffix.lower()] not in names:
            raise ValueError('File does not match its Office extension')
        if sum(i.file_size for i in zin.infolist()) > 1024**3:
            raise ValueError('Office package exceeds 1 GB uncompressed limit')
        removed = {n for n in names if n.lower().startswith('docprops/') and not n.lower().endswith('.xml')}
        with zipfile.ZipFile(destination, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                if name in removed:
                    continue
                data = zin.read(name)
                if name.lower().startswith('docprops/') and name.lower().endswith('.xml'):
                    root = ET.fromstring(data)
                    root.clear()
                    data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                elif name == '_rels/.rels':
                    root = ET.fromstring(data)
                    for child in list(root):
                        if child.attrib.get('Target', '').lstrip('/') in removed:
                            root.remove(child)
                    data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                elif name == '[Content_Types].xml':
                    root = ET.fromstring(data)
                    for child in list(root):
                        if child.attrib.get('PartName', '').lstrip('/') in removed:
                            root.remove(child)
                    data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                # New ZIP entries carry no source comments, extra fields or timestamps.
                item = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                item.compress_type = zipfile.ZIP_DEFLATED
                zout.writestr(item, data)
    with zipfile.ZipFile(destination) as check:
        if check.testzip():
            raise ValueError('Office archive verification failed')
        for name in check.namelist():
            if name.lower().startswith('docprops/') and name.lower().endswith('.xml'):
                root = ET.fromstring(check.read(name))
                if len(root) or root.attrib or (root.text or '').strip():
                    raise ValueError('Office property verification failed')


def empty_property_set(data):
    # OLE property-set header plus an empty section for each original FMTID.
    if len(data) < 28 or data[:2] != b'\xfe\xff':
        raise ValueError('Malformed legacy Office property stream')
    count = struct.unpack_from('<I', data, 24)[0]
    if count not in (1, 2) or len(data) < 28 + 28 * count:
        raise ValueError('Unsupported legacy property-set structure')
    result = bytearray(len(data))
    result[:2] = b'\xfe\xff'
    struct.pack_into('<I', result, 24, count)
    offset = 28 + 20 * count
    for i in range(count):
        descriptor = 28 + 20 * i
        result[descriptor:descriptor+16] = data[descriptor:descriptor+16]
        struct.pack_into('<I', result, descriptor+16, offset)
        struct.pack_into('<II', result, offset, 8, 0)
        offset += 8
    return bytes(result)


def clean_legacy(source, destination):
    import olefile
    shutil.copyfile(source, destination)
    with olefile.OleFileIO(destination, write_mode=True) as ole:
        expected = {'.doc': 'WordDocument', '.xls': 'Workbook', '.ppt': 'PowerPoint Document'}
        if not ole.exists(expected[source.suffix.lower()]):
            raise ValueError('Unrecognized legacy Office format')
        if ole.exists('EncryptedPackage') or ole.exists('EncryptionInfo'):
            raise ValueError('Encrypted Office document skipped')
        if source.suffix.lower() == '.doc':
            header = ole.openstream('WordDocument').read(12)
            if len(header) < 12 or struct.unpack_from('<H', header, 10)[0] & 0x8100:
                raise ValueError('Encrypted or obfuscated Word document skipped')
        for path in ole.listdir():
            if any('digitalsignature' in part.lower() for part in path):
                raise ValueError('Signed legacy Office document skipped')
            if path[-1] in ('\x05SummaryInformation', '\x05DocumentSummaryInformation'):
                data = ole.openstream(path).read()
                ole.write_stream(path, empty_property_set(data))
    with olefile.OleFileIO(destination) as check:
        for path in check.listdir():
            if path[-1] in ('\x05SummaryInformation', '\x05DocumentSummaryInformation'):
                data = check.openstream(path).read()
                if data != empty_property_set(data):
                    raise ValueError('Legacy property verification failed')


def clean_file(source, destination):
    source, destination = Path(source), Path(destination)
    suffix = source.suffix.lower()
    if suffix in EXTENSIONS:
        return clean_image(source, destination)
    if suffix not in DOCUMENTS:
        raise ValueError('Unsupported file type')
    destination.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix='.metadata-clean-', suffix=suffix, dir=destination.parent)
    os.close(fd)
    temporary = Path(temporary)
    try:
        if suffix == '.pdf':
            clean_pdf(source, temporary)
        elif suffix in {'.docx', '.xlsx', '.pptx'}:
            clean_office(source, temporary)
        else:
            clean_legacy(source, temporary)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)


def scan_files(root, recursive, excluded=None):
    result = []
    for directory, dirs, names in os.walk(root, followlinks=False, onerror=lambda e: (_ for _ in ()).throw(e)):
        dirs[:] = sorted(d for d in dirs if not d.startswith(('cleaned_images_', 'cleaned_files_'))
                         and not (Path(directory)/d).is_symlink()
                         and (excluded is None or (Path(directory)/d).resolve() != excluded)) if recursive else []
        result.extend(Path(directory)/n for n in sorted(names) if not (Path(directory)/n).is_symlink())
    return result


def export_report(rows, destination):
    with open(destination, 'w', newline='', encoding='utf-8-sig') as stream:
        writer = csv.DictWriter(stream, fieldnames=['file', 'status', 'detail'])
        writer.writeheader()
        # Prevent spreadsheet formula interpretation in exported filenames/errors.
        for row in rows:
            writer.writerow({key: ("'" + str(value) if str(value).startswith(('=', '+', '-', '@', '\t', '\r')) else value) for key, value in row.items()})


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('File Metadata Cleaner')
        self.geometry('860x730')
        self.minsize(760, 680)
        self.events = queue.Queue()
        self.cancel = threading.Event()
        self.busy = False
        self.rows = []
        self.folder = tk.StringVar()
        self.destination = tk.StringVar()
        self.recursive = tk.BooleanVar(value=False)
        self.mode = tk.StringVar(value='Folder')
        self.status = tk.StringVar(value='Select a source folder and export destination.')
        pane = ttk.Frame(self, padding=20)
        pane.pack(fill='both', expand=True)
        ttk.Label(pane, text='File Metadata Cleaner', font=('Segoe UI', 22, 'bold')).pack(anchor='w')
        ttk.Label(pane, text='Images • PDF • Word • Excel • PowerPoint').pack(anchor='w', pady=(4, 14))
        self.controls = []
        for label, var, command in [('Source folder', self.folder, self.select_folder), ('Export destination', self.destination, self.select_destination)]:
            ttk.Label(pane, text=label).pack(anchor='w')
            row = ttk.Frame(pane); row.pack(fill='x', pady=(3, 10))
            entry = ttk.Entry(row, textvariable=var); entry.pack(side='left', fill='x', expand=True)
            button = ttk.Button(row, text='Browse…', command=command); button.pack(side='left', padx=(8,0))
            self.controls.extend([entry, button])
        row = ttk.Frame(pane); row.pack(fill='x')
        ttk.Label(row, text='Export as:').pack(side='left')
        for label, value in [('Cleaned folder', 'Folder'), ('ZIP archive', 'ZIP'), ('Replace originals', 'Replace')]:
            button = ttk.Radiobutton(row, text=label, variable=self.mode, value=value, command=lambda: self.destination.set(''))
            button.pack(side='left', padx=8); self.controls.append(button)
        sub = ttk.Checkbutton(pane, text='Include subfolders', variable=self.recursive)
        sub.pack(anchor='w', pady=10); self.controls.append(sub)
        ttk.Label(pane, text='Removes image metadata and document properties. Keeps original file formats.\nDocument comments, revisions, embedded files and visible content are not scrubbed.\nJPEG is re-encoded. Signed PDFs/OOXML and animated images are skipped.\nLegacy DOC/XLS/PPT: standard property streams only; not a forensic scrub.', wraplength=810).pack(anchor='w', pady=(0,12))
        row = ttk.Frame(pane); row.pack(fill='x')
        start = ttk.Button(row, text='Clean and export', command=self.start)
        start.pack(side='left'); self.controls.append(start)
        self.stop = ttk.Button(row, text='Stop after current file', command=self.cancel.set, state='disabled')
        self.stop.pack(side='left', padx=8)
        self.report = ttk.Button(row, text='Export CSV report…', command=self.save_report, state='disabled')
        self.report.pack(side='right')
        self.progress = ttk.Progressbar(pane)
        self.progress.pack(fill='x', pady=(14,8))
        ttk.Label(pane, textvariable=self.status, wraplength=800).pack(anchor='w')
        frame = ttk.Frame(pane); frame.pack(fill='both', expand=True, pady=(8,0))
        self.log = tk.Text(frame, state='disabled', wrap='word', height=12)
        scroll = ttk.Scrollbar(frame, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y'); self.log.pack(fill='both', expand=True)
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.poll_timer = self.after(100, self.poll)

    def select_folder(self):
        chosen = filedialog.askdirectory(title='Source folder')
        if chosen: self.folder.set(chosen)

    def select_destination(self):
        if self.mode.get() == 'Replace':
            messagebox.showinfo('Replace originals', 'This mode saves each cleaned file over its original. Export destination is unused.')
            return
        chosen = (filedialog.asksaveasfilename(title='Export ZIP', defaultextension='.zip', filetypes=[('ZIP archive', '*.zip')])
                  if self.mode.get() == 'ZIP' else filedialog.askdirectory(title='Parent folder for a new cleaned_files folder'))
        if chosen: self.destination.set(chosen)

    def set_busy(self, busy):
        self.busy = busy
        for widget in self.controls: widget.configure(state='disabled' if busy else 'normal')
        self.stop.configure(state='normal' if busy else 'disabled')
        self.report.configure(state='normal' if not busy and self.rows else 'disabled')

    def start(self):
        root = Path(self.folder.get().strip()).expanduser().resolve()
        if not self.folder.get().strip() or not root.is_dir():
            messagebox.showerror('Source folder', 'Select an existing source folder.'); return
        mode = self.mode.get()
        destination = Path(self.destination.get().strip()).expanduser().resolve() if self.destination.get().strip() else root
        if mode == 'ZIP' and not self.destination.get().strip():
            messagebox.showerror('ZIP destination', 'Choose a ZIP export path.'); return
        if mode == 'ZIP' and (destination.suffix.lower() != '.zip' or not destination.parent.is_dir()):
            messagebox.showerror('ZIP destination', 'Choose a .zip path inside an existing folder.'); return
        if mode == 'ZIP' and destination.exists():
            messagebox.showerror('ZIP destination', 'Choose a new ZIP filename; existing exports are never overwritten.'); return
        if mode == 'Folder' and not destination.is_dir():
            messagebox.showerror('Export destination', 'Choose an existing parent folder.'); return
        if mode == 'Replace' and not messagebox.askyesno('Replace originals?', 'Permanently replace successfully cleaned source files? Keep your own backup. JPEG quality and color appearance may change.'):
            return
        self.rows = []; self.cancel.clear(); self.progress['value'] = 0
        self.log.configure(state='normal'); self.log.delete('1.0','end'); self.log.configure(state='disabled')
        self.status.set('Scanning…'); self.set_busy(True)
        threading.Thread(target=self.run_batch, args=(root, self.recursive.get(), mode, destination), daemon=True).start()

    def run_batch(self, root, recursive, mode, destination):
        rows = []
        staging = None
        try:
            files = scan_files(root, recursive, destination if mode == 'Folder' and destination != root else None)
            if not files:
                self.events.put(('done', (rows, 'No files found.'))); return
            if mode == 'Folder':
                output = Path(tempfile.mkdtemp(prefix='cleaned_files_', dir=destination))
            elif mode == 'ZIP':
                staging = tempfile.TemporaryDirectory(prefix='metadata-export-')
                output = Path(staging.name)
            else:
                output = root
            self.events.put(('total', len(files)))
            for index, source in enumerate(files, 1):
                if self.cancel.is_set(): break
                relative = source.relative_to(root)
                row = {'file': str(relative), 'status': 'cleaned', 'detail': 'Target metadata removed and output reopened for verification'}
                try:
                    if source.suffix.lower() not in ALL_EXTENSIONS:
                        row.update(status='skipped', detail='Unsupported file type')
                    else:
                        clean_file(source, output/relative)
                except Exception as error:
                    row.update(status='failed', detail=str(error))
                rows.append(row)
                self.events.put(('log', f'{row["status"]}: {relative} — {row["detail"]}'))
                self.events.put(('progress', index))
            processed = len(rows)
            for source in files[processed:]:
                rows.append({'file': str(source.relative_to(root)), 'status': 'unprocessed', 'detail': 'Cancelled'})
            if mode != 'Replace':
                # Pick a report filename that cannot replace a cleaned source file.
                report_path = output/'metadata_report.csv'
                while report_path.exists(): report_path = report_path.with_stem(report_path.stem + '_')
                export_report(rows, report_path)
            if mode == 'ZIP':
                try:
                    with open(destination, 'xb') as target:
                        with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as archive:
                            for path in output.rglob('*'):
                                if path.is_file(): archive.write(path, path.relative_to(output).as_posix())
                except FileExistsError:
                    raise ValueError('ZIP destination already exists; export not overwritten')
                except Exception:
                    destination.unlink(missing_ok=True)
                    raise
                with zipfile.ZipFile(destination) as check:
                    if check.testzip(): raise ValueError('ZIP export verification failed')
            target = destination if mode == 'ZIP' else output
            self.events.put(('log', f'Output: {target}'))
            counts = {status: sum(r['status'] == status for r in rows) for status in ('cleaned','skipped','failed','unprocessed')}
            label = 'Stopped; partial export' if self.cancel.is_set() else 'Finished'
            self.events.put(('done', (rows, label + ': ' + ', '.join(f'{v} {k}' for k,v in counts.items()))))
        except Exception as error:
            self.events.put(('done', (rows, f'Batch/export failed: {error}')))
        finally:
            if staging: staging.cleanup()

    def save_report(self):
        path = filedialog.asksaveasfilename(title='Export results', defaultextension='.csv', filetypes=[('CSV report','*.csv')])
        if path:
            try: export_report(self.rows, path)
            except Exception as error: messagebox.showerror('Report export failed', str(error))

    def poll(self):
        if getattr(self, "poll_timer", None):
            self.after_cancel(self.poll_timer)
            self.poll_timer = None
        for _ in range(100):
            try: kind, value = self.events.get_nowait()
            except queue.Empty: break
            if kind == 'log':
                self.log.configure(state='normal'); self.log.insert('end', value+'\n'); self.log.see('end'); self.log.configure(state='disabled')
            elif kind == 'total': self.progress['maximum'] = value
            elif kind == 'progress':
                self.progress['value'] = value
                self.status.set(f'Processed {value} of {int(self.progress["maximum"])} files…')
            elif kind == 'done':
                self.rows, status = value
                self.status.set(status); self.set_busy(False)
        self.poll_timer = self.after(100, self.poll)

    def destroy(self):
        if getattr(self, "poll_timer", None):
            self.after_cancel(self.poll_timer)
            self.poll_timer = None
        super().destroy()

    def close(self):
        if self.busy:
            self.cancel.set(); self.status.set('Stopping after current file and finishing export. Close when finished.')
        else: self.destroy()


if __name__ == '__main__':
    App().mainloop()
