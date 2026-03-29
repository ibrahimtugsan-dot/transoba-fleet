# 🚢 Fleet Noon Report — Kurulum Rehberi

## Nasıl Çalışır?

```
Gemi → Excel e-mail → Outlook → Power Automate → GitHub /data/ → Actions → index.html → GitHub Pages
```

**Tamamen otomatik. Sıfır manuel adım.**

---

## Adım 1 — GitHub Repo Kur (5 dk)

1. [github.com/new](https://github.com/new) → repo adı: `vessel_reports_monitoring_transoba`
2. **Public** seç (GitHub Pages ücretsiz için)
3. Bu repo'nun tüm dosyalarını yükle:
   - `template.html`
   - `generate_report.py`
   - `.github/workflows/update_report.yml`
   - `data/` (boş klasör — `.gitkeep` dosyası ekle)

### GitHub Pages'i Aç
`Settings → Pages → Source: Deploy from branch → Branch: main → / (root)`

URL'in: `https://KULLANICI_ADIN.github.io/vessel_reports_monitoring_transoba/`

---

## Adım 2 — GitHub Personal Access Token (2 dk)

Power Automate'in dosya yükleyebilmesi için token gerekiyor.

1. GitHub → `Settings → Developer settings → Personal access tokens → Tokens (classic)`
2. `Generate new token (classic)`
3. İsim: `vessel_reports_monitoring_transoba-bot`
4. Expiration: `No expiration`
5. Scope: ✅ `repo` (sadece bu)
6. **Token'ı kopyala ve kaydet** (bir daha göremezsin!)

---

## Adım 3 — Power Automate Flow (10 dk)

[make.powerautomate.com](https://make.powerautomate.com) → `Create → Automated cloud flow`

### Tetikleyici: "When a new email arrives (V3)"
```
Folder:     Inbox
From:       kaptan@gemi.com  ← gemilerden gelen e-mail adresleri
Has Attachments: Yes
```
> Birden fazla gemi için "From" alanını boş bırak, aşağıda filtrele.

---

### Action 1 — "Apply to each" (attachments)
```
Select an output: Attachments
```

**İçinde: Condition**
```
Attachment Name — ends with — .xlsx
```

---

### Action 2 — "HTTP" (Condition True içinde)
```
Method:  PUT
URI:     https://api.github.com/repos/KULLANICI/vessel_reports_monitoring_transoba/contents/data/@{items('Apply_to_each')?['name']}

Headers:
  Authorization:  token BURAYA_TOKEN_YAPISTIR
  Content-Type:   application/json
  Accept:         application/vnd.github+json

Body:
{
  "message": "📎 @{items('Apply_to_each')?['name']} eklendi",
  "content": "@{base64(body('Get_attachment_(V2)'))}",
  "sha": "@{if(empty(outputs('Get_existing_file')?['body']?['sha']), '', outputs('Get_existing_file')?['body']?['sha'])}"
}
```

> **Not:** Aynı isimde dosya varsa SHA gerekiyor. Bunun için önce bir "GET" action ekle:

### Action 2a — "HTTP" (mevcut dosyayı kontrol et)
```
Method:  GET
URI:     https://api.github.com/repos/KULLANICI/vessel_reports_monitoring_transoba/contents/data/@{items('Apply_to_each')?['name']}
Headers:
  Authorization:  token BURAYA_TOKEN_YAPISTIR
  Accept:         application/vnd.github+json
```
`Name: Get_existing_file` — hata verse de devam et (Configure run after → Has failed ✅)

---

### Dosya İsimlendirme — Hiçbir Kural Yok ✅

`generate_report.py` dosya adını **tamamen yok sayar**.  
Gemi adını, rapor tipini ve tarihi doğrudan Excel içinden okur.

- `noon_28032026_v2_FINAL.xlsx` → Olur ✓
- `OCEAN DESTINY March Report.xlsx` → Olur ✓
- `rapor kopyası (3).xlsx` → Olur ✓

**Aynı gemi için birden fazla dosya** varsa (kaptanlar tekrar gönderdiyse)  
script en son tarihli veriyi otomatik seçer, eskisini görmezden gelir.

Power Automate'te **rename yapmana gerek yok** — dosyayı olduğu gibi yükle.

---

## Adım 4 — Test

1. Bir test xlsx'i Outlook'a e-mail olarak gönder
2. Power Automate `My flows → Run history` → başarılı mı?
3. GitHub repo → `data/` klasöründe dosya görünüyor mu?
4. `Actions` sekmesi → workflow çalıştı mı?
5. 1-2 dk bekle → Pages URL'ini aç

---

## Alternatif: Outlook Klasör Kuralı

Kaptanlar farklı konular / adresler kullanıyorsa:

1. Outlook → `Rules` → yeni kural
2. **Koşul:** "Has attachment" + "Subject contains: NOON REPORT"
3. **Eylem:** "Move to folder: Noon Reports"

Power Automate'te Folder'ı `Noon Reports` yap.

---

## Sorun Giderme

| Sorun | Çözüm |
|-------|-------|
| Actions çalışmıyor | `Settings → Actions → Allow all actions` |
| Pages açılmıyor | `Settings → Pages` → branch seçildi mi? |
| SHA hatası | GET action "Continue on failure" ayarla |
| Grafik görünmüyor | xlsx dosyasında sheet adı NOON REPORT veya PORT REPORT olmalı |
| Veri gelmiyor | `generate_report.py` satır 300+ hata loguna bak |

---

## Sonuç

Her sabah kaptanlar raporu e-mail ile gönderir →  
**5 dakika içinde** `https://KULLANICI.github.io/vessel_reports_monitoring_transoba/` güncellenir.  
Ekip sadece URL'i bookmark'lar. Başka hiçbir şey gerekmez.
