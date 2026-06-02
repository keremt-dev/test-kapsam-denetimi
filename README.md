# test-kapsam-denetimi

Test senaryolarının kullanım senaryolarını ne kadar karşıladığını **madde madde** denetleyen ve kapsam boşluğu raporu üreten bir [Claude Code skill](https://docs.claude.com/en/docs/claude-code/skills)'i. TCDD BVA (BVAKP) modülleri (KEY, AT, …) için tasarlandı; aynı doküman şablonunu kullanan her modüle uygulanabilir.

## Ne yapar?

Üç kaynak dokümandan yola çıkar:

| Kaynak | Açıklama |
|--------|----------|
| **SRS gereksinim** `.xlsx` | `Gereksinim Numarası` başlıklı gereksinim listesi |
| **Kullanım senaryoları** `.docx` | Her KS bir tablo (ana akış, alternatifler, iş kuralları, ön/son koşul) |
| **Test senaryoları** `.xlsx` | Her KS bir sayfa; ID, adım, beklenen sonuç, gereksinim etiketi |

Her kullanım senaryosunu **ayrı bir paralel ajan (Opus 4.8)** ile denetler. Her KS şu 5 boyutta atomik maddelere ayrılıp test senaryolarıyla **Tam / Kısmi / Yok** olarak eşleştirilir:

- **Ana Akış** adımları
- **Alternatif Senaryolar**
- **İş Kuralları**
- **İlgili SRS Gereksinimleri**
- **Ön / Son Koşullar**

Ters yönde de denetler: hangi testin hangi maddeyi karşıladığı, gereksinim etiketi boş testler, ve yalnızca başlık içeren artık/iskelet test satırları.

## Çıktılar

- `<MOD>_Kapsam_Bosluk_Raporu.md` — yönetici özeti, kritik boşluklar, KS bazında detay
- `<MOD>_Kapsam_Bosluk_Matrisi.xlsx` — renk kodlu izlenebilirlik matrisi (`Matris` / `Ozet` / `Etiketsiz_Artik`)
- `<test dosyası> - Inceleme.xlsx` — orijinalin kopyası + her satıra `Kapsam İnceleme Notu` sütunu (orijinale dokunulmaz)
- `<MOD>_jira_yorumlari.md` — KS başına insansı/emojisiz, eklemeye hazır Jira yorumları

## Kurulum

```bash
git clone https://github.com/keremt-dev/test-kapsam-denetimi.git \
  ~/.claude/skills/test-kapsam-denetimi
pip install openpyxl python-docx
```

Claude Code skill'i otomatik keşfeder. Tetikleyiciler: *"test kapsam denetimi"*, *"kapsam boşluğu"*, *"coverage gap"*, *"test senaryoları yeterli mi"* …

## Kullanım

Bir modülün üç kaynak dosyasını bir klasöre koyup skill'i tetikleyin. Manuel akış:

```bash
SKILL_DIR=~/.claude/skills/test-kapsam-denetimi

# Faz 0 — çıkarım
python "$SKILL_DIR/scripts/extract_bundles.py" --indir . --outdir work

# Faz 1 — her KS için paralel Opus 4.8 ajanı (Claude Code orkestre eder)
#          work/<KS>_input.md oku → work/<KS>_result.json yaz

# Faz 3 — sert doğrulama (PASS olmadan sentez yok)
python "$SKILL_DIR/scripts/verify.py" --workdir work

# Faz 4 — sentez
python "$SKILL_DIR/scripts/synthesize.py" --workdir work --outdir .
python "$SKILL_DIR/scripts/verify.py" --workdir work --post --outdir .
```

> Windows'ta tüm komutları `PYTHONIOENCODING=utf-8` ile çalıştırın.

## Doğrulama (test kontrol yapısı)

`verify.py`, sentezden **önce** çalışan sert bir kapıdır; herhangi bir FAIL'de teslim yapılmaz:

| Kontrol | Amaç |
|---------|------|
| **C1** | Sayımları madde durumundan **yeniden hesaplar**; ajanın `ozet` değerine güvenilmez |
| **C2** | Atıf yapılan her Test ID kaynak dosyada gerçekten var (uydurma yok) |
| **C3** | Her geçerli Test ID, test notlarında tam bir kez geçer |
| **C4** | KS'nin SRS boyutu, beklenen gereksinim listesiyle birebir |
| **C5** | Tüm KS'lerin SRS birleşimi = modülün SRS evreni (eksik/fazla yok) |
| **C6** | `durum` ve `tur` alanları geçerli enum |

`--post` modu sentez sonrası matris satır sayısı ve özet toplamlarını doğrular.

## Tasarım ilkeleri

- **Tek doğruluk kaynağı = madde `durum` alanı.** Hiçbir çıktı ajan özet sayımına güvenmez.
- **Başlık-tabanlı tespit** — kolon/tablo sırası değişse de çalışır.
- **Kapsam-içi oran** — Jira oranlarında düşük öncelikli ön/son koşul kısmileri paydadan düşülür.

## Yapı

```
test-kapsam-denetimi/
├── SKILL.md
├── README.md
├── scripts/
│   ├── extract_bundles.py   # Faz 0: kaynaklar → girdi paketleri + manifest
│   ├── verify.py            # Faz 3: sert doğrulama (C1–C6, --post)
│   └── synthesize.py        # Faz 4: rapor + matris + yorumlu kopya + jira
└── references/
    ├── agent_prompt_template.md
    └── result_schema.json
```

## Lisans

MIT
