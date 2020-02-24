from pathlib import Path
from typing import Any, Optional


def _get_first_file_in_dir(dir: str) -> Optional[str]:
    dir_p = Path(dir)
    if not dir_p.is_dir():
        return None

    for content in dir_p.iterdir():
        if content.is_file():
            return str(content)

    return None


def _init_dir(dir: str, unlink: bool) -> Optional[str]:
    path = Path(dir)

    # 指定されたディレクトリと同名のファイルがあったらエラー
    if path.is_file():
        return None

    # 指定されたディレクトリが存在しなかったら作成
    if not path.exists():
        path.mkdir(parents=True)
        return str(path)

    # 既存のディレクトリで、unlink = True だったら中身を全削除
    if unlink:
        __deep_rmdir(path)

    return str(path.resolve())


def __deep_rmdir(path: Path):
    if path.is_dir():
        for content in path.iterdir():
            if content.is_file():
                content.unlink()
            elif content.is_dir():
                __deep_rmdir(content)
                content.rmdir()
