from pathlib import Path
from typing import Optional, List
from PyPDF2 import PdfFileMerger

from shipping_instruction.util import _init_dir


def merge(inputDir: str,
          outputBaseDir: str,
          instructionNumber: str) -> Optional[str]:
    output = __init_output_dir(outputBaseDir, instructionNumber)
    if output is None:
        return None

    merger = PdfFileMerger()
    counter = 0
    for pdf in __get_original_files(inputDir):
        merger.append(pdf)
        counter += 1

    if counter == 0:
        return None

    o = Path(output)
    if not o.suffix == ".pdf":
        return None

    # 出力する前に同名のファイルがあったらエラー
    if o.exists():
        return None

    merger.write(output)
    merger.close()

    if not o.is_file():
        return None

    return output


def __init_output_dir(base: str, name: str) -> Optional[str]:
    b = _init_dir(base, False)
    if b is None:
        return None

    d = Path(b).joinpath(name)
    if _init_dir(str(d), True) is None:
        return None

    return str(d.joinpath(f"{name}.pdf"))


def __get_original_files(dir: str) -> List[str]:
    p = Path(dir)
    if not p.is_dir():
        return []

    files = []
    for content in p.iterdir():
        if not content.is_file():
            continue

        if not content.suffix == ".pdf":
            continue

        files.append(str(content))

    return files
