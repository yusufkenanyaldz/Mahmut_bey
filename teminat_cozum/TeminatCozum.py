"""Windows .exe giriş noktası (PyInstaller).

Argümansız çalıştırılırsa (çift tıklama) pencereli arayüz açılır; bir klasör .exe'nin üzerine sürüklenirse pencere o
klasör seçili açılır. Diğer argümanlarla komut satırıyla aynıdır:
    TeminatCozum.exe taslak "<firmanın o ayki klasörü>" --ayar firmalar\\<firma>\\ayar.yaml
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
    if len(sys.argv) == 2 and Path(sys.argv[1]).is_dir():
        # klasör .exe'nin üzerine sürüklenip bırakıldı: pencereyi o klasör seçili açar
        from teminat.arayuz import main as arayuz_main
        arayuz_main(sys.argv[1])
        return 0
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
