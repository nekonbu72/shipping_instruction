import shutil
import subprocess
from pathlib import Path

import PyInstaller.__main__


def build():
    PyInstaller.__main__.run([
        "--noconfirm",
        "--log-level=WARN",
        "--onefile",
        "--clean",
        "shipping_instruction\\main.py"
    ])


def rm_cache():
    shutil.rmtree("build", True)


def rm_spec():
    p = Path("main.spec")
    if p.is_file():
        p.unlink()


def clean_build():
    build()
    rm_cache()
    rm_spec()


def run_exe():
    subprocess.run("dist\\main.exe")
