# Qaiməni cədvələ çevir

Windows üçün kiçik Python tətbiqi. Şəkil kompüterdə oxunur. Nəticəni pəncərədə düzəldə və Excel, JSON və ya tablarla ayrılmış mətn kimi saxlaya bilərsiniz.

## Bu kompüterdə işə salmaq

`Qaimeni-ac.cmd` faylını açın və ya layihə qovluğunda `python app.py` işlədin. OCR mühərriki `.ocr-runtime` qovluğunda, Azərbaycan və ingilis dili faylları isə `tessdata` qovluğundadır.

## Başqa kompüterə köçürmək

Windows-da 64-bit Python 3.12 ilə layihə qovluğunda bu əmrləri işlədin. Layihənin `tessdata` qovluğuna [Tesseract dil paketlərindən](https://github.com/tesseract-ocr/tessdata_fast) `aze.traineddata` və `eng.traineddata` fayllarını əlavə edin.

```powershell
python -m pip install -r requirements.txt
python -m pip install --target .ocr-runtime tesserocr==2.13.0
python app.py
```

## İstifadə

**Şəkil seç** düyməsi ilə qaimənin məhsul cədvəli görünən PNG və ya JPG faylını açın. Nəticədə sətirlər göründükdə lazım olan düzəlişləri edin. Sütunları **Tab** düyməsi ilə ayırmaq olar. Sonra formatı seçib faylı saxlayın.

**Avtomatik** rejim əvvəlcə ingilis modeli ilə nömrələnmiş məhsul sətirlərini axtarır. Qaimə sxemi tanınmadıqda Azərbaycan dili modelinə keçir. Hər iki dili ayrıca seçmək də mümkündür. Nümunədəki qaimə sxemi tanındıqda yalnız məhsul cədvəli göstərilir və ixrac edilir.

Proqram sözlərin şəkildəki mövqeyinə əsasən sətir və sütunları təxmin edir. Boş xanalar, çoxsətirli məhsul adları və əyri şəkillər sütunları sürüşdürə bilər. Excel xanaları mətn kimi saxlanır ki, məhsul kodlarının əvvəlindəki sıfırlar və onluq vergüllər qorunsun. JSON faylında `rows` açarı hər sətir üçün bir massiv saxlayır.

## Ollama ilə ayrıca sınaq

`Ollama-test.ps1` JPG və ya PNG şəklini lokal Ollama-ya göndərir və nəticəni PowerShell-də JSON kimi göstərir. Şəkli əvvəlcə məhsul cədvəlinə qədər kəsin ki, bank və əlaqə məlumatları modelə ötürülməsin. Sınaq üçün Ollama işləməli və `qwen3.5:4b` modeli serverdə yüklənmiş olmalıdır.

```powershell
.\Ollama-test.ps1 -ImagePath "C:\Qaimeler\mehsullar.jpg"
```

Başqa şəkil dəstəkli model üçün `-Model` parametrini verin. Cavabı orijinal qaimə ilə tutuşdurun. Tesseract tətbiqinin açılışı və Excel ixracı əvvəlki qaydada işləyir.
