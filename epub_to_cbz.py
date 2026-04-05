"""
epub_to_cbz.py
将 EPUB 漫画文件转换为 CBZ 格式，严格保持页面顺序。
打包命令：pyinstaller --onefile --console epub_to_cbz.py
"""

import os
import sys
import zipfile
import shutil
import re
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET


# ── 常量 ──────────────────────────────────────────────────────────────────────
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tiff", ".avif"}
NS = {
    "opf": "http://www.idpf.org/2007/opf",
    "dc":  "http://purl.org/dc/elements/1.1/",
    "cnt": "urn:oasis:names:tc:opendocument:xmlns:container",
}


# ── 核心逻辑 ──────────────────────────────────────────────────────────────────
def find_opf_path(zf: zipfile.ZipFile) -> str:
    """从 META-INF/container.xml 找到 OPF 文件路径。"""
    try:
        container_xml = zf.read("META-INF/container.xml")
        root = ET.fromstring(container_xml)
        rootfile = root.find(".//cnt:rootfile", NS)
        if rootfile is not None:
            return rootfile.attrib.get("full-path", "")
    except Exception:
        pass
    # 回退：直接搜索 .opf
    for name in zf.namelist():
        if name.endswith(".opf"):
            return name
    return ""


def get_spine_image_paths(zf: zipfile.ZipFile, opf_path: str) -> list[str]:
    """
    解析 OPF 文件，按 <spine> 顺序返回图片路径列表。
    支持两种常见漫画 EPUB 结构：
      1. spine 指向图片文件（直接）
      2. spine 指向 XHTML，XHTML 内嵌 <img>（间接）
    """
    opf_dir = str(Path(opf_path).parent)
    if opf_dir == ".":
        opf_dir = ""

    opf_xml = zf.read(opf_path)
    root = ET.fromstring(opf_xml)

    # 建立 manifest id → href 映射
    manifest: dict[str, str] = {}
    for item in root.findall(".//opf:manifest/opf:item", NS):
        item_id   = item.attrib.get("id", "")
        item_href = item.attrib.get("href", "")
        media     = item.attrib.get("media-type", "")
        # 规范化路径（相对于 OPF 所在目录）
        full = normalize_path(opf_dir, item_href)
        manifest[item_id] = (full, media)

    # 按 spine 顺序取 idref
    spine_items = root.findall(".//opf:spine/opf:itemref", NS)
    ordered_images: list[str] = []

    for itemref in spine_items:
        idref = itemref.attrib.get("idref", "")
        if idref not in manifest:
            continue
        full_path, media_type = manifest[idref]
        ext = Path(full_path).suffix.lower()

        if ext in IMAGE_EXTS or "image" in media_type:
            # 直接是图片
            ordered_images.append(full_path)
        elif ext in (".xhtml", ".html", ".htm") or "html" in media_type:
            # 从 XHTML 中提取 <img src="...">
            imgs = extract_images_from_xhtml(zf, full_path, opf_dir)
            ordered_images.extend(imgs)

    return ordered_images


def extract_images_from_xhtml(zf: zipfile.ZipFile, xhtml_path: str, opf_dir: str) -> list[str]:
    """从单个 XHTML 页面中按顺序提取图片路径。"""
    results = []
    try:
        content = zf.read(xhtml_path)
        # 去除命名空间前缀以兼容各种写法
        content = re.sub(rb'\s+xmlns[^=]*="[^"]*"', b"", content)
        content = re.sub(rb'<(\w+):(\w+)', rb'<\2', content)
        content = re.sub(rb'</(\w+):(\w+)', rb'</\2', content)
        root = ET.fromstring(content)
        xhtml_dir = str(Path(xhtml_path).parent)

        for img in root.iter("img"):
            src = img.attrib.get("src", "") or img.attrib.get("xlink:href", "")
            if src:
                full = normalize_path(xhtml_dir, src)
                if Path(full).suffix.lower() in IMAGE_EXTS:
                    results.append(full)
        # SVG image 标签
        for image in root.iter("image"):
            src = image.attrib.get("{http://www.w3.org/1999/xlink}href", "") or image.attrib.get("href", "")
            if src:
                full = normalize_path(xhtml_dir, src)
                if Path(full).suffix.lower() in IMAGE_EXTS:
                    results.append(full)
    except Exception as e:
        print(f"  [警告] 解析 {xhtml_path} 时出错: {e}")
    return results


def normalize_path(base_dir: str, href: str) -> str:
    """将相对 href 拼合为 ZIP 内绝对路径（不含前导斜杠）。"""
    # 去掉 fragment
    href = href.split("#")[0]
    if href.startswith("/"):
        return href.lstrip("/")
    if base_dir:
        combined = base_dir + "/" + href
    else:
        combined = href
    # 解析 ../ ./
    parts = []
    for part in combined.split("/"):
        if part == "..":
            if parts:
                parts.pop()
        elif part not in (".", ""):
            parts.append(part)
    return "/".join(parts)


def epub_to_cbz(epub_path: str, output_dir: str | None = None) -> str:
    """
    主转换函数。
    返回输出的 CBZ 文件完整路径。
    """
    epub_path = Path(epub_path).resolve()
    if not epub_path.exists():
        raise FileNotFoundError(f"找不到文件：{epub_path}")
    if epub_path.suffix.lower() != ".epub":
        raise ValueError(f"不是 EPUB 文件：{epub_path}")

    out_dir = Path(output_dir) if output_dir else epub_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    cbz_path = out_dir / (epub_path.stem + ".cbz")

    print(f"\n📖 读取：{epub_path.name}")

    with zipfile.ZipFile(epub_path, "r") as zf:
        all_names = set(zf.namelist())

        # 1. 定位 OPF
        opf_path = find_opf_path(zf)
        if not opf_path:
            raise RuntimeError("无法找到 OPF 文件，可能不是标准 EPUB。")
        print(f"   OPF：{opf_path}")

        # 2. 按顺序获取图片路径
        image_paths = get_spine_image_paths(zf, opf_path)

        # 3. 去重（保留首次出现顺序）
        seen: set[str] = set()
        unique_images: list[str] = []
        for p in image_paths:
            if p not in seen and p in all_names:
                seen.add(p)
                unique_images.append(p)

        # 4. 如果 spine 方法未找到图片，回退到按文件名排序
        if not unique_images:
            print("   [回退] Spine 中无图片，改用文件名排序……")
            all_images = sorted(
                [n for n in all_names if Path(n).suffix.lower() in IMAGE_EXTS],
                key=lambda x: natural_sort_key(x)
            )
            unique_images = all_images

        if not unique_images:
            raise RuntimeError("EPUB 中未找到任何图片！")

        print(f"   找到图片：{len(unique_images)} 页")

        # 5. 写入 CBZ
        digits = len(str(len(unique_images)))
        with zipfile.ZipFile(cbz_path, "w", zipfile.ZIP_DEFLATED) as cbz:
            for idx, img_path in enumerate(unique_images, start=1):
                ext = Path(img_path).suffix.lower()
                new_name = f"{idx:0{digits}d}{ext}"
                data = zf.read(img_path)
                cbz.writestr(new_name, data)
                print(f"   [{idx:>{digits}}/{len(unique_images)}] {Path(img_path).name} → {new_name}")

    print(f"\n✅ 完成！输出：{cbz_path}\n")
    return str(cbz_path)


def natural_sort_key(s: str):
    """自然排序键，使 page2 < page10。"""
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r"(\d+)", s)]


# ── CLI 入口 ───────────────────────────────────────────────────────────────────
def main():
    args = sys.argv[1:]

    # 拖拽多个文件或命令行传参均支持
    if not args:
        print("=" * 55)
        print("  EPUB → CBZ 漫画转换工具")
        print("=" * 55)
        print("用法：")
        print("  1. 直接将 .epub 文件拖拽到本程序图标上")
        print("  2. 命令行：epub_to_cbz.exe file1.epub [file2.epub ...]")
        print("  3. 命令行（指定输出目录）：")
        print("     epub_to_cbz.exe --output D:\\comics file.epub")
        print()
        # 交互模式
        path = input("请输入 EPUB 文件路径（或直接拖入）：").strip().strip('"')
        if not path:
            sys.exit(0)
        args = [path]

    # 解析 --output 参数
    output_dir = None
    filtered = []
    i = 0
    while i < len(args):
        if args[i] in ("--output", "-o") and i + 1 < len(args):
            output_dir = args[i + 1]
            i += 2
        else:
            filtered.append(args[i])
            i += 1

    success, failed = 0, []
    for epub in filtered:
        try:
            epub_to_cbz(epub.strip().strip('"'), output_dir)
            success += 1
        except Exception as e:
            print(f"\n❌ 转换失败 [{epub}]：{e}\n")
            failed.append(epub)

    print(f"─── 汇总：成功 {success} 个，失败 {len(failed)} 个 ───")
    if failed:
        for f in failed:
            print(f"  ✗ {f}")

    input("\n按 Enter 键退出……")


if __name__ == "__main__":
    main()
