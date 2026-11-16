"""rapor_sablonu.html ve Chart.js'i yerel/bddk_indir.py içine gömer (gzip + base64).

Şablon ya da Chart.js değiştiğinde çalıştırın:  python scripts/yerel_paketle.py
"""
import base64
import gzip
import re
from pathlib import Path

KOK = Path(__file__).resolve().parent.parent
HEDEF = KOK / "yerel" / "bddk_indir.py"


def gom(yol):
    return base64.b64encode(gzip.compress(yol.read_bytes(), 9, mtime=0)).decode()


s = HEDEF.read_text("utf-8")
for ad, etiket, kaynak in [("SABLON_GZ_B64", "<<SABLON>>", KOK / "rapor_sablonu.html"),
                           ("CHARTJS_GZ_B64", "<<CHARTJS>>", KOK / "yerel" / "vendor" / "chart.umd.min.js")]:
    s, n = re.subn(rf'^{ad} = ".*"  # {etiket}$', f'{ad} = "{gom(kaynak)}"  # {etiket}', s, flags=re.M)
    assert n == 1, ad
HEDEF.write_text(s, "utf-8")
print(f"{HEDEF}: {HEDEF.stat().st_size / 1e3:.0f} KB")
