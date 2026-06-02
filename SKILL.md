---
name: test-kapsam-denetimi
description: Test senaryolarinin kullanim senaryolarini ne kadar karsiladigini madde madde denetler ve kapsam boslugu raporu uretir. Uc kaynak dokumandan (SRS gereksinim .xlsx, kullanim senaryolari .docx, test senaryolari .xlsx) yola cikip her kullanim senaryosunu ayri bir paralel ajan (Opus 4.8) ile denetler; ana akis adimlari, alternatif senaryolar, is kurallari, on/son kosullar ve SRS gereksinimlerini Tam/Kismi/Yok olarak eslestirir. Cikti olarak kapsam boslugu raporu (.md), izlenebilirlik matrisi (.xlsx), yorumlu test kopyasi (.xlsx) ve KS bazinda Jira yorumlari (.md) uretir; sonunda sert bir dogrulama (test kontrol) yapisi calistirir. Tetikleyiciler - "test kapsam denetimi", "kapsam boslugu", "coverage gap", "test kullanim senaryosu kapsam", "test senaryolari yeterli mi", "use case test coverage", "BVAKP test kapsam", "test eksik bul", "kapsam matrisi".
---

# test-kapsam-denetimi

Test senaryolarinin kullanim senaryolarini madde madde ne kadar karsiladigini denetleyip kapsam boslugu raporlari ureten arac. BVAKP modulleri (KEY, AT, ...) icin tasarlandi; ayni sablonu kullanan her modul/dosya setine uygulanabilir.

## Triggers
- "test kapsam denetimi" / "kapsam boslugu" / "coverage gap"
- "test senaryolari kullanim senaryolarini karsiliyor mu"
- "test eksik bul" / "kapsam matrisi" / "use case test coverage"

## Girdi
Bir klasor (veya 3 dosya yolu) gerekir:
1. **SRS gereksinim .xlsx** — `Gereksinim Numarasi` basligi olan dosya.
2. **Kullanim senaryolari .docx** — her KS bir tablo (Kullanim Senaryo Numarasi, Ilgili Gereksinimler, On/Son Kosul, Basarili Ana Senaryo, Alternatif Senaryolar, Is Kurallari).
3. **Test senaryolari .xlsx** — her KS bir sayfa (KS_001, KS_002, ...); kolonlar: ID, Test Senaryo Adi, Aciklama, On Kosul, Gereksinim, Test Adimlari, Beklenen Sonuc.

## Ortam Ayarlari
Script ve referanslar bu skill dizinindedir. Calisma basinda `SKILL_DIR` belirle (ilk bulunan):
1. Bu dosyanin dizini
2. `~/.claude/skills/test-kapsam-denetimi/`

```bash
for d in "$HOME/.claude/skills/test-kapsam-denetimi" "./test-kapsam-denetimi"; do
  [ -f "$d/scripts/extract_bundles.py" ] && echo "SKILL_DIR=$d" && break
done
```
Windows'ta tum python cagrilarinda `PYTHONIOENCODING=utf-8` kullan.

## Dependencies
- Python 3.10+, `openpyxl >= 3.1`, `python-docx >= 1.1`
```bash
pip install openpyxl python-docx
```

## Scripts
| Script | Amac |
|--------|------|
| `scripts/extract_bundles.py` | Faz 0 — 3 kaynaktan `work/<KS>_input.md` + `work/manifest.json` uretir (baslik-tabanli tespit) |
| `scripts/verify.py` | Faz 3 — sert dogrulama (PASS/FAIL); ozet'i yeniden hesaplar; `--post` ile cikti kontrolu |
| `scripts/synthesize.py` | Faz 4 — rapor.md + matris.xlsx + yorumlu kopya + jira.md (sayimlar madde durumundan) |
| `references/agent_prompt_template.md` | KS denetci ajan promptu |
| `references/result_schema.json` | result.json semasi |

---

## İŞ AKIŞI (sirayla uygula)

### Faz 0 — Cikarim
Calisma dizinine gec (girdi dosyalarinin oldugu yer onerilir) ve calistir:
```bash
PYTHONIOENCODING=utf-8 python "$SKILL_DIR/scripts/extract_bundles.py" --indir . --outdir work
```
Bu, `work/` altinda her KS icin `<KS>_input.md` ve `manifest.json` uretir. Ciktidaki KS sayisi, modul kodu ve SRS sayisini kullaniciya bildir. (Dosyalari acikca belirtmek istersen `--reqs/--usecases/--tests` kullan.)

### Faz 1 — Paralel denetim ajanlari (Opus 4.8)
`manifest.json`'daki HER kullanim senaryosu icin **tek bir mesajda, paralel** Agent cagrisi yap:
- `subagent_type: "general-purpose"`
- **`model: "opus"`** (Opus 4.8 — bu ZORUNLU; her KS ajani Opus 4.8 ile calismali)
- prompt: `references/agent_prompt_template.md` icerigini al, `{KS}` `{INPUT_PATH}` (`work/<KS>_input.md`) ve `{RESULT_PATH}` (`work/<KS>_result.json`) yer tutucularini doldur.

Her ajan kendi `work/<KS>_result.json` dosyasini yazar ve kisa ozet doner. KS'ler bagimsizdir; hepsini ayni anda baslat.

### Faz 3 — SERT DOGRULAMA (sentezden ONCE kapi)
```bash
PYTHONIOENCODING=utf-8 python "$SKILL_DIR/scripts/verify.py" --workdir work
```
- C1 sayim butunlugu (ozet'i madde durumundan yeniden hesaplar/yazar) · C2 uydurma Test ID yok · C3 her test icin test_notu var · C4 KS SRS-boyutu == beklenen · C5 modul SRS butunlugu · C6 enum.
- **PASS degilse sentez YAPMA.** FAIL'lerde: ilgili KS ajanini DUZELTILMIS yonergeyle yeniden calistir (orn. eksik SRS maddesi ekle, uydurma ID'yi kaldir), sonra `verify.py`'yi tekrar kosur. PASS olana dek tekrarla.

### Faz 4 — Sentez (yalniz PASS sonrasi)
```bash
PYTHONIOENCODING=utf-8 python "$SKILL_DIR/scripts/synthesize.py" --workdir work --outdir .
```
Uretir: `<MOD>_Kapsam_Bosluk_Raporu.md`, `<MOD>_Kapsam_Bosluk_Matrisi.xlsx`, `<test dosyasi> - Inceleme.xlsx`, `<MOD>_jira_yorumlari.md`. Dosya kilitliyse (Excel acik) uyarir; kullaniciya kapatmasini soyleyip yeniden calistir.

Son olarak cikti kontrolu:
```bash
PYTHONIOENCODING=utf-8 python "$SKILL_DIR/scripts/verify.py" --workdir work --post --outdir .
```

## Cikti dosyalari
- `<MOD>_Kapsam_Bosluk_Raporu.md` — yonetici ozeti (KS tablosu, %), Yok maddeler, Kismiler, ters yon, KS detaylari.
- `<MOD>_Kapsam_Bosluk_Matrisi.xlsx` — `Matris` (renk kodlu), `Ozet`, `Etiketsiz_Artik`.
- `<test dosyasi> - Inceleme.xlsx` — orijinalin kopyasi + her sayfada `Kapsam Inceleme Notu` sutunu (orijinale dokunulmaz).
- `<MOD>_jira_yorumlari.md` — KS basina insansi/emojisiz; kapsam-ici oran; yalniz Yok + onemli Kismi.

## Tasarim notlari (degistirme)
- **Sayimlar HER ZAMAN `maddeler[].durum`'dan** hesaplanir; ajanin `ozet` alanina guvenilmez (gecmiste 15/16 sayim hatasina yol acti).
- Kolon/tablo tespiti **baslik adina** gore; sira indeksine guvenilmez.
- Jira oranlarinda dusuk oncelikli On/Son Kosul Kismileri paydadan dusulur ("ele alinan N madde").
