"""Create a reproducible source handoff without credentials or user run data."""
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parents[1]
TOP_FILES = {
    'app.py', 'core.py', 'planner.py', 'universal.py', 'universal_api.py',
    'sandbox.py', 'desktop.py', 'TableBridge.spec', 'README.md',
    'requirements.txt', 'requirements-desktop.txt', '.env.example', '.gitignore',
    '启动.command', '启动.bat', '桌面启动.bat', '构建Windows.bat',
}
SOURCE_DIRS = {'static', 'samples', 'scripts', 'tests', 'docs'}


def package(destination):
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    files = []
    for path in sorted(ROOT.rglob('*')):
        if not path.is_file():
            continue
        rel = path.relative_to(ROOT)
        if any(part in {'__pycache__', '.venv', 'data', 'build', 'dist', '.git'} for part in rel.parts):
            continue
        if path.name in {'.env', '.DS_Store'} or path.suffix in {'.pyc', '.zip'}:
            continue
        if len(rel.parts) == 1 and rel.name in TOP_FILES or rel.parts[0] in SOURCE_DIRS:
            files.append(path)
    manifest = {}
    with ZipFile(destination, 'w', ZIP_DEFLATED) as z:
        for path in files:
            name = 'tablebridge/' + path.relative_to(ROOT).as_posix()
            raw = path.read_bytes()
            z.writestr(name, raw)
            manifest[name] = hashlib.sha256(raw).hexdigest()
        z.writestr('SHA256SUMS.json', json.dumps(manifest, ensure_ascii=False, indent=2))
    with ZipFile(destination) as z:
        assert z.testzip() is None
        assert 'tablebridge/.env' not in z.namelist()
        assert not any('/data/' in name or '/.venv/' in name for name in z.namelist())
    return {'path': str(destination.resolve()), 'source_files': len(files), 'bytes': destination.stat().st_size, 'sha256': hashlib.sha256(destination.read_bytes()).hexdigest()}


if __name__ == '__main__':
    import sys
    if len(sys.argv) != 2:
        raise SystemExit('Usage: python scripts/package_source.py <output.zip>')
    print(json.dumps(package(sys.argv[1]), ensure_ascii=False))
