from pathlib import Path
import zipfile

src = Path('__10_.zip')
out = Path('test-data')
out.mkdir(exist_ok=True)
with zipfile.ZipFile(src) as archive:
    for info in archive.infolist():
        name = info.filename
        try:
            name = name.encode('cp437').decode('gbk')
        except UnicodeError:
            pass
        target = out / name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(archive.read(info))
print(f'extracted {len(archive.infolist())} files to {out}')
