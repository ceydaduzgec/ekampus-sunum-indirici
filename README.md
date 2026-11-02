# Slayt İndirici ve Sunum Oluşturucu

Bu Python kodu, belirtilen bir URL'den SVG slaytlarını indirir, JPEG veya PNG formatına dönüştürür ve bir PowerPoint sunumu oluşturur.

## Kurulum Talimatları

1. **Gerekli Bağımlılıkları Yükleyin**

   Öncelikle, sisteminizde Python 3.11+ yüklü olduğundan emin olun. Ardından gerekli paketleri yükleyin:

   ```bash
   python3 -m venv env
   source env/bin/activate
   pip install -r requirements.txt
   ```

   Not: CairoSVG'nin düzgün çalışması için, ek sistem bağımlılıklarını yüklemeniz gerekebilir:

   - **Ubuntu/Debian'da**:
     ```bash
     sudo apt-get install libcairo2-dev libpango1.0-dev libgdk-pixbuf2.0-dev libffi-dev shared-mime-info
     ```

   - **macOS'ta** (Homebrew kullanarak):
     ```bash
     brew install cairo pango gdk-pixbuf libffi
     ```

   - **Windows'ta**:
     pip'in otomatik olarak yükleyeceği önceden derlenmiş ikili paketleri kullanmanız önerilir.

## 2. Slaytın URL'sini Alın

Tarayıcınızda bir slayta sağ tıklayın ve slaytın URL'sini almak için "Bağlantı Adresini Kopyala" veya "Bağlantıyı Kopyala" seçeneğini seçin.

<div align="center">
  <img src="images/url_kopyalama.png" alt="Slayt URL'sini kopyalama" width="80%" style="border: 1px solid #ddd; border-radius: 5px;" />
</div>

## 3. Kodu Çalıştırın

   Slaytları indirmek ve bir sunum oluşturmak için kodu kullanın:

   ```bash
   python3 slide_downloader.py https://vkscalelite.anadolu.edu.tr/presentation/633fc11f5c4832a9db123d72eb67831a012d0710-1741272327834/presentation/091f41e4a0d9eaa5c28fad9596121aa1ed2b315c-1741272345267/svgs/slide1.svg
   ```

   Kod şunları yapacaktır:
   - Tüm slaytları indirir (slide1.svg'den başlayarak)
   - Bunları JPEG formatına dönüştürür
   - Tüm slaytlarla bir PowerPoint sunumu oluşturur
   - Geçici görüntü dosyalarını otomatik olarak siler

## Komut Satırı Seçenekleri

```
Kullanım: slide_downloader.py [-h] [--input FILE] [--output-dir OUTPUT_DIR] [--start START]
                          [--max MAX] [--format {jpeg,png}] [--presentation PRESENTATION]
                          [--keep-images]
                          [base_url]

Slaytları indirir ve bir sunum oluşturur

Konumsal Değişkenler:
  base_url              Slaytlar için temel URL (örn., https://example.com/presentation/id/svgs/slide)
                        (--input kullanılmıyorsa zorunludur)

Opsiyonel Değişkenler:
  -h, --help                   Yardım mesajını göster ve çık
  --input FILE                 Her satırı "Sunum adı: URL" olan metin dosyası; base_url yerine
                               her satır için ayrı bir "<ad>.pptx" oluşturur
  --output-dir OUTPUT_DIR      Görüntülerin kaydedileceği dizin (varsayılan: slides)
  --start START                Başlangıç slayt numarası (varsayılan: 1)
  --max MAX                    İndirilecek maksimum slayt sayısı (varsayılan: 100)
  --format {jpeg,png}          Çıktı görüntü formatı (varsayılan: jpeg)
  --presentation PRESENTATION  Çıktı PowerPoint dosyası (varsayılan: presentation.pptx); --input ile yok sayılır
  --keep-images                Sunum oluşturulduktan sonra geçici görüntü dosyalarını sakla
```

## Örnekler

1. **Temel kullanım**:
   ```bash
   python slide_downloader.py https://vkscalelite.anadolu.edu.tr/presentation/633fc11f5c4832a9db123d72eb67831a012d0710-1741272327834/presentation/091f41e4a0d9eaa5c28fad9596121aa1ed2b315c-1741272345267/svgs/slide
   ```

2. **Slayt aralığını belirtme**:
   ```bash
   python slide_downloader.py https://vkscalelite.anadolu.edu.tr/presentation/633fc11f5c4832a9db123d72eb67831a012d0710-1741272327834/presentation/091f41e4a0d9eaa5c28fad9596121aa1ed2b315c-1741272345267/svgs/slide --start 5 --max 20
   ```

3. **PNG formatını kullanma**:
   ```bash
   python slide_downloader.py https://vkscalelite.anadolu.edu.tr/presentation/633fc11f5c4832a9db123d72eb67831a012d0710-1741272327834/presentation/091f41e4a0d9eaa5c28fad9596121aa1ed2b315c-1741272345267/svgs/slide --format png
   ```

4. **Özel çıktı konumu**:
   ```bash
   python slide_downloader.py https://vkscalelite.anadolu.edu.tr/presentation/633fc11f5c4832a9db123d72eb67831a012d0710-1741272327834/presentation/091f41e4a0d9eaa5c28fad9596121aa1ed2b315c-1741272345267/svgs/slide --output-dir slaytlarim --presentation dersim.pptx
   ```

5. **Geçici görüntü dosyalarını sakla**:
   ```bash
   python slide_downloader.py https://vkscalelite.anadolu.edu.tr/presentation/633fc11f5c4832a9db123d72eb67831a012d0710-1741272327834/presentation/091f41e4a0d9eaa5c28fad9596121aa1ed2b315c-1741272345267/svgs/slide --keep-images
   ```

6. **Birden fazla haftayı tek seferde indirme (`--input`)**:
   ```bash
   python3 slide_downloader.py --input haftalar.txt
   ```
   Ayrıntılar için aşağıdaki **Girdi Dosyası ile Toplu İndirme** bölümüne bakın.

## Girdi Dosyası ile Toplu İndirme

Her hafta için komutu ayrı ayrı çalıştırmak yerine, tüm haftaları bir `.txt` dosyasına yazıp `--input` ile tek seferde indirebilirsiniz. Dosyadaki her satır **bir haftanın sunumudur** ve `Sunum adı: URL` biçimindedir. URL, o haftanın sunumunun başlangıç linkidir; yani tek sunum indirirken `base_url` olarak verdiğiniz link.

```text
2024-2025 Bahar_Hafta 1: https://vkscalelite.anadolu.edu.tr/presentation/.../svgs/slide1.svg
2024-2025 Bahar_Hafta 2: https://vkscalelite.anadolu.edu.tr/presentation/.../svgs/slide1.svg
```

```bash
python3 slide_downloader.py --input haftalar.txt
```

Bu komut her satır için mevcut komutu ayrı ayrı çalıştırmışsınız gibi davranır. Örneğin ilk satır şu komuta karşılık gelir:

```bash
python3 slide_downloader.py https://vkscalelite.anadolu.edu.tr/presentation/633fc11f5c4832a9db123d72eb67831a012d0710-1741272327834/presentation/091f41e4a0d9eaa5c28fad9596121aa1ed2b315c-1741272345267/svgs/slide1.svg --presentation "2024-2025 Bahar_Hafta 1.pptx"
```

- Her satır için, satırdaki adla bir sunum oluşturulur (`2024-2025 Bahar_Hafta 1.pptx` gibi) ve komutun çalıştırıldığı klasöre kaydedilir. Dosya adında kullanılamayan karakterler (`/ \ : * ? " < > |`) `_` ile değiştirilir.
- Satırlar dosyadaki sırayla işlenir.
- Ad ile URL, satırdaki **ilk** `: ` (iki nokta + boşluk) ile ayrılır. `https://` içindeki iki nokta ayırıcı sayılmaz. Bu yüzden sunum adının içinde `: ` kullanmayın.
- Boş satırlar atlanır. Dosya UTF-8 olmalıdır. Aynı ad iki kez kullanılamaz.
- `--start`, `--max`, `--format` ve `--keep-images` her sunuma ayrı ayrı uygulanır. `--keep-images` verilirse görüntüler `--output-dir` altında her sunum için ayrı bir klasörde saklanır. `--presentation` bu modda yok sayılır.
- `base_url` ve `--input` birlikte kullanılamaz.
- Dosya okunamazsa ya da bir satırın biçimi hatalıysa, indirme başlamadan önce hatalı satır numaralarıyla birlikte bir hata mesajı gösterilir.
- Bir hafta indirilemezse diğer haftalar indirilmeye devam eder. İşlem sonunda başarısız satırlar listelenir.

## Testler

Testler slaytları yerel bir HTTP sunucusundan sunar, bu yüzden internet bağlantısı gerekmez. Kurulumdaki bağımlılıklar (CairoSVG'nin sistem bağımlılıkları dahil) yüklü olmalıdır. Proje klasöründe şu komutu çalıştırın:

```bash
python3 -m unittest discover tests
```

## Sorun Giderme

1. **SVG Dönüştürme Sorunları**:
   CairoSVG bazı slaytları dönüştüremiyorsa, bu hatalı biçimlendirilmiş SVG içeriğinden kaynaklanıyor olabilir. SVG'leri ayrı ayrı indirmeyi ve farklı bir dönüştürücü kullanmayı deneyin.

2. **404 Hataları**:
   Kod, mevcut slaytların sonuna ulaştığını varsayarak 3 ardışık 404 hatasından sonra otomatik olarak duracaktır.

3. **Hız Sınırlaması**:
   Bağlantı hataları alıyorsanız, koddaki `time.sleep()` değerini değiştirerek indirmeler arasındaki gecikmeyi artırmayı deneyin.