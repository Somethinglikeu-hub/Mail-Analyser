# Mail-Analyser

Outlook maillerini PC'de metne çevirir, GitHub'a atar; GitHub Actions bunları OpenRouter üzerinden ucuz bir LLM ile analiz edip **DASHBOARD.md** ve günlük raporlar üretir.

```
PC (Windows, Outlook)                        GitHub (private repo)
collector/mail_collector.py  ── git push ──▶ feed/YYYY-MM-DD/mails/*.txt
  (17:00 görev / collect_now.bat)                    │  push tetikler
                                              .github/workflows/analyze.yml
                                                     │  analyzer/analyze.py ──▶ OpenRouter
                                                     ▼
                                              reports/YYYY-MM-DD.md + .json
                                              DASHBOARD.md  ◀── arayüz burası
```

- **collector/** — PC'de çalışır. Outlook'u sadece okur, mailleri + eklerin metnini `feed/`'e yazar, sonra `git push` yapar. Orijinal ekler (`feed/*/files/`) ve `state.json` PC'de kalır, GitHub'a gitmez.
- **analyzer/** — GitHub Actions'ta çalışır. Sadece yeni mailleri analiz eder (önceki sonuçlar `reports/*.json`'da cache'li), her mail için özet / öncelik / aksiyon / son tarih / cevap beklenen çıkarır, günlük brifing yazar.
- **analyzer/prompts/** — LLM talimatları. Ne çıkarılacağını değiştirmek için kod değil bu dosyalar düzenlenir.

## Kurulum

### 1. Repo'yu private yap (zorunlu)
Settings → General → Danger Zone → **Change visibility → Private**.
Public repoda collector push etmeyi, workflow da çalışmayı reddeder — iş mailleri herkese açık olur.

### 2. OpenRouter anahtarı
1. https://openrouter.ai → Keys → yeni key, Credits'ten birkaç $ yükle.
2. Repo → Settings → Secrets and variables → Actions:
   - **Secrets** → `OPENROUTER_API_KEY` = key
   - **Variables** (opsiyonel) → `OPENROUTER_MODEL` (varsayılan `google/gemini-2.5-flash`), `MAIL_OWNER` (ör. `Kerem`)

Model listesi ve fiyatlar: https://openrouter.ai/models. Her çalışmanın gerçek maliyeti Actions logunun son satırında yazar.

### 3. PC tarafı (bir kere)
```bat
git clone https://github.com/Somethinglikeu-hub/Mail-Analyser C:\MailAnalyser
cd C:\MailAnalyser
pip install -r collector\requirements.txt
```
- Eski klasördeki `state.json`'ı `collector\` içine kopyala (aynı mailler tekrar toplanmasın).
- İstersen eski `feed\` günlerini repo kökündeki `feed\` içine kopyala; ilk push'ta onlar da analiz edilir.
- `git push`'un şifre sormadan çalıştığını bir kere elle dene (Git Credential Manager ile giriş yap).
- Zamanlanmış görevi yeni yere taşı: `powershell -ExecutionPolicy Bypass -File collector\fix_task.ps1`

## Kullanım
- Her gün 17:00'de görev kendiliğinden çalışır; anında istersen `collector\collect_now.bat`.
- 1-2 dk sonra repoda **DASHBOARD.md** güncellenir (telefonda GitHub uygulamasından da açılır).
- Tekrar analiz: Actions → *Analyze mails* → Run workflow (gün ve `force` seçilebilir).
- Hata olursa: PC'de `collector\run_log.txt` (`PUSH ...` satırı), GitHub'da Actions logu.

## Yerelde deneme
```bash
export OPENROUTER_API_KEY=sk-or-...
python analyzer/analyze.py --day 2026-10-23
```

## Gizlilik
Mail metinleri GitHub'a (private repo) ve OpenRouter üzerinden seçilen modelin sağlayıcısına gider. Şirket politikası buna izin vermiyorsa kullanma; OpenRouter ayarlarında "no training / zero data retention" sağlayıcılarla sınırlayabilirsin.
