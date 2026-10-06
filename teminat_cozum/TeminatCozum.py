"""Windows .exe giriş noktası (PyInstaller).

Argümansız çalıştırılırsa (çift tıklama) pencereli arayüz açılır. Argümanla çalıştırılırsa komut satırıyla aynıdır:
    TeminatCozum.exe taslak --ayar firmalar\\<firma>\\ayar.yaml --kok "...\\<FİRMA>" --donem 2026-03
Pencereli .exe'nin konsolu olmadığı için bu durumda çıktı .exe'nin yanındaki (yazılamıyorsa TEMP klasöründeki)
teminat_son_calisma.log dosyasına yazılır.
"""
import os
import sys
import tempfile
from pathlib import Path

GUNLUK = 'teminat_son_calisma.log'


def _gunluk_ac():
    for klasor in (Path(sys.executable).parent, Path(tempfile.gettempdir())):
        try:
            return open(klasor / GUNLUK, 'w', encoding='utf-8')
        except OSError:
            continue
    return open(os.devnull, 'w', encoding='utf-8')


def main():
    if len(sys.argv) > 1:
        if sys.stdout is None or sys.stderr is None:
            sys.stdout = sys.stderr = _gunluk_ac()
        from teminat.cli import main as cli_main
        kod = cli_main()
        sys.stdout.flush()
        return kod
    from teminat.arayuz import main as arayuz_main
    arayuz_main()
    return 0


if __name__ == '__main__':
    sys.exit(main())
