""".doc → .docx dönüşümü (LibreOffice). Orijinal dosyaya dokunulmaz; çıktı ayrı klasöre yazılır."""
import os
import shutil
import subprocess
import tempfile
from pathlib import Path


class DonusumHatasi(RuntimeError):
    pass


def soffice_yolu():
    aday = [os.environ.get('SOFFICE'), shutil.which('soffice'), shutil.which('soffice.exe'), shutil.which('libreoffice')]
    for kok in (os.environ.get('PROGRAMFILES'), os.environ.get('PROGRAMFILES(X86)'), r'C:\Program Files', r'C:\Program Files (x86)'):
        if kok:
            aday.append(str(Path(kok) / 'LibreOffice' / 'program' / 'soffice.exe'))
    for a in aday:
        if a and Path(a).is_file():
            return a
    return None


def docx_hazirla(path, onbellek):
    """.docx ise yolu aynen döndürür; .doc ise `onbellek` klasörüne .docx'e çevirip yeni yolu döndürür.

    Kaynaktan daha yeni bir dönüşüm zaten varsa yeniden çevrilmez.
    """
    path = Path(path)
    if path.suffix.lower() == '.docx':
        return path
    if path.suffix.lower() != '.doc':
        raise DonusumHatasi(f'Word dosyası değil: {path}')
    onbellek = Path(onbellek)
    onbellek.mkdir(parents=True, exist_ok=True)
    hedef = onbellek / (path.stem + '.docx')
    if hedef.exists() and hedef.stat().st_mtime >= path.stat().st_mtime:
        return hedef
    so = soffice_yolu()
    if not so:
        raise DonusumHatasi('LibreOffice (soffice) bulunamadı; .doc dosyasını Word ile .docx olarak kaydedin '
                            'ya da SOFFICE ortam değişkenine soffice.exe yolunu yazın.')
    # Açık bir LibreOffice penceresi dönüşümü sessizce engellemesin diye ayrı profil kullanılır.
    with tempfile.TemporaryDirectory() as profil:
        komut = [so, f'-env:UserInstallation={Path(profil).as_uri()}', '--headless',
                 '--convert-to', 'docx', '--outdir', str(onbellek), str(path)]
        try:
            subprocess.run(komut, check=True, capture_output=True, timeout=180)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as e:
            raise DonusumHatasi(f'Dönüştürülemedi: {path.name} ({e})') from e
    if not hedef.exists():
        raise DonusumHatasi(f'Dönüştürülemedi: {path.name} (LibreOffice çıktı üretmedi)')
    return hedef
