"""Compila excel_merger.py a un .exe optimizado con PyInstaller.

Uso:
    python build_exe.py
    python build_exe.py --onedir
    python build_exe.py --console
    python build_exe.py --upx-dir "C:\\ruta\\a\\upx"

Por defecto genera un único archivo:
    dist/CombinadorExcel.exe
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
PROGRAMA = RAIZ / "excel_merger.py"
NOMBRE = "CombinadorExcel"
DIST = RAIZ / "dist"
BUILD = RAIZ / "build"

# Módulos pesados que el programa no usa. Excluirlos reduce el tamaño.
EXCLUIR = [
    "jupyter",
    "jupyterlab",
    "notebook",
    "ipykernel",
    "IPython",
    "matplotlib",
    "scipy",
    "PyQt5",
    "PyQt6",
    "PySide2",
    "PySide6",
    "pytest",
    "sphinx",
    "docutils",
]


def verificar_dependencias() -> None:
    """Comprueba que el Python que ejecuta el script tenga lo necesario."""
    faltantes = []
    for modulo in ("pandas", "openpyxl", "xlrd", "tkinter"):
        try:
            __import__(modulo)
        except ImportError:
            faltantes.append(modulo)

    if faltantes:
        print("Faltan dependencias en este Python:")
        print("  " + ", ".join(faltantes))
        print()
        print("Instalalas con:")
        print(f'  "{sys.executable}" -m pip install -r requirements.txt')
        raise SystemExit(1)


def asegurar_pyinstaller() -> None:
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("PyInstaller no está instalado. Instalando...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "pyinstaller"]
        )


def construir(args: argparse.Namespace) -> Path:
    comando = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--name",
        NOMBRE,
        "--distpath",
        str(DIST),
        "--workpath",
        str(BUILD),
        "--specpath",
        str(BUILD),
        # Quita docstrings y asserts del bytecode empaquetado.
        "--optimize",
        "2",
        "--hidden-import",
        "openpyxl",
        "--hidden-import",
        "xlrd",
    ]

    comando.append("--onedir" if args.onedir else "--onefile")
    comando.append("--console" if args.console else "--windowed")

    for modulo in EXCLUIR:
        comando.extend(["--exclude-module", modulo])

    if args.upx_dir:
        comando.extend(["--upx-dir", str(Path(args.upx_dir).resolve())])

    comando.append(str(PROGRAMA))

    print("Compilando", PROGRAMA.name, "...")
    print(" ".join(f'"{parte}"' if " " in parte else parte for parte in comando))
    print()
    subprocess.check_call(comando, cwd=RAIZ)

    if args.onedir:
        return DIST / NOMBRE / f"{NOMBRE}.exe"
    return DIST / f"{NOMBRE}.exe"


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Compila excel_merger.py a un .exe con PyInstaller."
    )
    parser.add_argument(
        "--onedir",
        action="store_true",
        help="Genera una carpeta con el .exe (arranca más rápido).",
    )
    parser.add_argument(
        "--console",
        action="store_true",
        help="Muestra consola al abrir el .exe (útil para depurar).",
    )
    parser.add_argument(
        "--upx-dir",
        help="Carpeta de UPX para comprimir el ejecutable (opcional).",
    )
    args = parser.parse_args()

    if not PROGRAMA.exists():
        raise SystemExit(f"No encontré {PROGRAMA}")

    verificar_dependencias()
    asegurar_pyinstaller()
    ejecutable = construir(args)

    if ejecutable.exists():
        tamano = ejecutable.stat().st_size / (1024 * 1024)
        print()
        print(f"Listo: {ejecutable}  ({tamano:.1f} MB)")
    else:
        print()
        print("La compilación terminó, pero no encontré el .exe esperado:")
        print(f"  {ejecutable}")


if __name__ == "__main__":
    main()
