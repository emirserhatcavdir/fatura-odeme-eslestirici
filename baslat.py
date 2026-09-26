"""baslat.bat için yalnızca yerel sunucu başlatma ve port kontrolü."""

import argparse
from importlib.util import find_spec
from pathlib import Path
import socket
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
HOST = "127.0.0.1"


def port_number(value: str) -> int:
    try:
        result = int(value)
    except ValueError:
        raise argparse.ArgumentTypeError("Port bir tam sayı olmalı; örnek: 8502.") from None
    if not 1 <= result <= 65535:
        raise argparse.ArgumentTypeError("Port 1 ile 65535 arasında olmalı.")
    return result


def port_available(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as listener:
        if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
            listener.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        try:
            listener.bind((HOST, port))
        except OSError:
            return False
    return True


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Fatura-Ödeme Eşleştirici yerel başlatıcı")
    parser.add_argument("port", nargs="?", type=port_number, default=8501, help="İsteğe bağlı port; varsayılan: 8501")
    args = parser.parse_args(argv)
    if not (ROOT / "sap.py").is_file():
        print("sap.py bulunamadı. Başlatıcıyı proje dosyalarıyla aynı klasörde tutun.")
        return 1
    if find_spec("streamlit") is None:
        print("Projenin sanal ortamında Streamlit bulunamadı.")
        print("README.md içindeki kurulum adımlarını tamamlayın; otomatik paket kurulumu yapılmaz.")
        return 1
    if not port_available(args.port):
        print(f"{args.port} portu kullanımda veya erişime kapalı. Hiçbir süreç sonlandırılmadı.")
        print(f"Uygulama zaten çalışıyorsa mevcut adresi kullanın: http://{HOST}:{args.port}")
        alternative = 8502 if args.port != 8502 else 8503
        print(f"Başka boş bir port seçmek için proje klasöründe çalıştırın: .\\baslat.bat {alternative}")
        return 1
    command = [sys.executable, "-m", "streamlit", "run", str(ROOT / "sap.py"),
               "--server.address", HOST, "--server.port", str(args.port), "--server.headless", "true"]
    print(f"Uygulama başlatılıyor: http://{HOST}:{args.port}", flush=True)
    print("Bu pencereyi açık tutun. Durdurmak için Ctrl+C kullanın.", flush=True)
    try:
        return subprocess.run(command, cwd=ROOT, check=False).returncode
    except KeyboardInterrupt:
        print("\nBaşlatıcı kapatıldı.")
        return 0
    except OSError as exc:
        print(f"Sunucu başlatılamadı: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
