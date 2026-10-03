"""Rapor çıktıları (GUI'den bağımsız)."""
import re

import pandas as pd


def export_sections(file_path, sections):
    """OrderedDict(başlık → DataFrame) yapısını çok sayfalı Excel'e yazar."""
    used = set()
    with pd.ExcelWriter(file_path) as writer:
        for title, df in sections.items():
            name = re.sub(r"[\[\]:*?/\\]", "", title)[:31] or "Sayfa"
            base, n = name, 2
            while name in used:
                suffix = f" {n}"
                name = base[:31 - len(suffix)] + suffix
                n += 1
            used.add(name)
            df.to_excel(writer, sheet_name=name, index=False)
