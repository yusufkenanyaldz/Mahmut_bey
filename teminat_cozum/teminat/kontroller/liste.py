"""Kontrol listesi: her çalıştırmada üretilen bulgular (rapora değil, ayrı listeye)."""
from dataclasses import dataclass, field

HATA, UYARI, ELLE, BILGI, TAMAM = 'HATA', 'UYARI', 'ELLE', 'BİLGİ', 'TAMAM'
DURUMLAR = (HATA, UYARI, ELLE, BILGI, TAMAM)
DURUM_ACIKLAMA = {
    HATA: 'Taslak bu haliyle yanlış olabilir; mutlaka bakılmalı.',
    UYARI: 'Tutarsızlık veya takdir gerektiren seçim; kontrol edilmeli.',
    ELLE: 'Girdi dosyalarında olmayan bilgi; elle doldurulacak (taslakta sarı).',
    BILGI: 'Bilgi amaçlı not.',
    TAMAM: 'Kontrol yapıldı, sorun yok.',
}


@dataclass
class Kontrol:
    durum: str
    konu: str
    aciklama: str
    no: str = ''          # PROJE_TALIMATI.md bölüm 6'daki madde numarası (varsa)
    yer: str = ''         # raporda yeri: 'Tablo 3-4-3', 'Paragraf 4-5' ...
    beklenen: str = ''
    bulunan: str = ''


@dataclass
class KontrolListesi:
    kalemler: list = field(default_factory=list)

    def ekle(self, durum, konu, aciklama, **kw):
        assert durum in DURUMLAR, durum
        k = Kontrol(durum, konu, aciklama, **kw)
        self.kalemler.append(k)
        return k

    def hata(self, konu, aciklama, **kw):
        return self.ekle(HATA, konu, aciklama, **kw)

    def uyari(self, konu, aciklama, **kw):
        return self.ekle(UYARI, konu, aciklama, **kw)

    def elle(self, konu, aciklama, **kw):
        return self.ekle(ELLE, konu, aciklama, **kw)

    def bilgi(self, konu, aciklama, **kw):
        return self.ekle(BILGI, konu, aciklama, **kw)

    def tamam(self, konu, aciklama, **kw):
        return self.ekle(TAMAM, konu, aciklama, **kw)

    def sonuc(self, kosul, konu, tamam_metni, hata_metni, durum=HATA, **kw):
        """kosul doğruysa TAMAM, değilse verilen durumla kayıt."""
        return self.ekle(TAMAM if kosul else durum, konu, tamam_metni if kosul else hata_metni, **kw)

    def sirali(self):
        sira = {d: i for i, d in enumerate(DURUMLAR)}
        return sorted(self.kalemler, key=lambda k: sira[k.durum])

    def say(self, durum):
        return sum(1 for k in self.kalemler if k.durum == durum)

    def ozet(self):
        return {d: self.say(d) for d in DURUMLAR}

    def __iter__(self):
        return iter(self.kalemler)

    def __len__(self):
        return len(self.kalemler)
