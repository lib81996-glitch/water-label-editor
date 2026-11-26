"""Package a Windows build and its documentation into a portable ZIP."""
import sys
from importlib.metadata import distribution
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

root = Path(__file__).resolve().parents[1]
name = "水洗标签编辑器"
exe = root / "dist" / f"{name}.exe"
archive_path = root / "dist" / f"{name}-v1.12-便携分发版.zip"
files = (
    root / "LICENSE", root / "使用说明.txt", root / "运行环境检测.bat",
    root / "标准底版-35x75.wslabel", root / "示例-6608.wslabel",
    *sorted(path for path in (root / "assets").rglob("*") if path.is_file()),
)
if not exe.is_file():
    raise SystemExit(f"请先构建 {exe.name}")

with ZipFile(archive_path, "w", compression=ZIP_DEFLATED, compresslevel=6) as archive:
    archive.write(exe, f"{name}/{exe.name}")
    for path in files:
        archive.write(path, f"{name}/{path.relative_to(root).as_posix()}")
    third_party = [("Python", Path(sys.base_prefix) / "LICENSE.txt")]
    for library, filename in (("Pillow", "LICENSE"), ("PyInstaller", "COPYING.txt")):
        package = distribution(library)
        notice = next(path for path in package.files
                      if path.name == filename and path.parent.name == "licenses")
        third_party.append((library, package.locate_file(notice)))
    for library, path in third_party:
        archive.write(path, f"{name}/第三方许可证/{library}.txt")

with ZipFile(archive_path) as archive:
    if archive.testzip() is not None:
        raise RuntimeError("分发包校验失败")
print(archive_path)
