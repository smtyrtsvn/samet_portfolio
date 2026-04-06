import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import barcode
from barcode.writer import ImageWriter
import os


def create_professional_sticker(data_dict, filename="etiket.png"):
    width, height = 410, 540  # Yükseklik aşağı kaymalardan dolayı hafif artırıldı
    img = Image.new('RGB', (width, height), color='white')
    draw = ImageDraw.Draw(img)

    try:
        draw.rounded_rectangle([15, 15, width - 15, height - 15], radius=45, outline="red", width=2)
    except AttributeError:
        draw.rectangle([15, 15, width - 15, height - 15], outline="red", width=2)

    # --- FONT BOYUTLARI GÜNCELLENDİ ---
    try:
        font_logo = ImageFont.truetype("times.ttf", 40, )  # WILLIAMS SONOMA (Sabit) "timesbd.ttf"
        font_sanfran = ImageFont.truetype("arial.ttf", 28)  # %25 Büyüdü (22 -> 28) "arialbd.ttf"
        font_urun = ImageFont.truetype("arial.ttf", 28)  # %20 Büyüdü (22 -> 26)
        font_made_in = ImageFont.truetype("arial.ttf", 28)  # Sabit
        font_sku = ImageFont.truetype("arial.ttf", 42)  # %15 Büyüdü (32 -> 37)
        font_side = ImageFont.truetype("arial.ttf", 42)  # %15 Büyüdü (40 -> 46)
        font_price = ImageFont.truetype("arial.ttf", 42)  # %15 Büyüdü (50 -> 58)
    except IOError:
        font_logo = font_sanfran = font_urun = font_made_in = font_sku = font_side = font_price = ImageFont.load_default()

    # --- MUTLAK Y KONUMLARI (1mm = ~12 piksel hesabı) ---
    y_williams = 23  # Sabit
    y_sonoma = 67  #
    y_sanfran = 115  # 110 aşağı
    y_urun = 155  # 160 yukarı
    y_made_in = 200  #
    y_barcode = 230  # 4mm aşağı (Eski: 186 -> Yeni: 234)
    y_sku = 325  # 330
    y_side = 365  # 3mm aşağı (Eski: 336 -> Yeni: 372)
    y_line = 422  # Orantılı olarak aşağı alındı
    y_price = 450  # Orantılı olarak aşağı alındı

    # 1. Brand Logo
    if data_dict["brand"] == "WS":
        draw.text(((width - draw.textlength("WILLIAMS", font=font_logo)) / 2, y_williams), "WILLIAMS", font=font_logo,
                  fill="black")
        draw.text(((width - draw.textlength("SONOMA", font=font_logo)) / 2, y_sonoma), "SONOMA", font=font_logo,
                  fill="black")
    elif data_dict["brand"] == "PBT":
        # Logoyu yükle
        pb_logo = Image.open("pottery_barn_teen_logo.png")

        # İhtiyaca göre yeniden boyutlandır (Örn: 250px genişlik, 80px yükseklik)
        pb_logo = pb_logo.resize((250, 80), Image.Resampling.LANCZOS)

        # Logoyu x ve y koordinatlarına yapıştır (ortalamak için x'i hesaplıyoruz)
        logo_x = (width - 250) // 2
        logo_y = 25
        img.paste(pb_logo, (logo_x, logo_y))
    elif data_dict["brand"] == "PBK":
        # Logoyu yükle
        pb_logo = Image.open("pottery_barn_kids_logo.png")

        # İhtiyaca göre yeniden boyutlandır (Örn: 250px genişlik, 80px yükseklik)
        pb_logo = pb_logo.resize((250, 80), Image.Resampling.LANCZOS)

        # Logoyu x ve y koordinatlarına yapıştır (ortalamak için x'i hesaplıyoruz)
        logo_x = (width - 250) // 2
        logo_y = 25
        img.paste(pb_logo, (logo_x, logo_y))
    # 2. Adres, Ürün ve Made In
    draw.text(((width - draw.textlength("San Francisco, CA 94109", font=font_sanfran)) / 2, y_sanfran),
              "San Francisco, CA 94109", font=font_sanfran, fill="black")
    draw.text(((width - draw.textlength(data_dict["urun_adi"], font=font_urun)) / 2, y_urun), data_dict["urun_adi"],
              font=font_urun, fill="black")
    draw.text(((width - draw.textlength("MADE IN TURKEY", font=font_made_in)) / 2, y_made_in), "MADE IN TURKEY",
              font=font_made_in, fill="black")

    # 3. Barkod (Boyut %10 Ufaltıldı)
    options = {
        'write_text': False,
        'module_width': 0.35,
        'module_height': 15.0,
        'quiet_zone': 0.5
    }

    code_128 = barcode.get('code128', data_dict['sku'], writer=ImageWriter())
    temp_filename = code_128.save('temp_barcode', options=options)
    barcode_img = Image.open(temp_filename)
    barcode_w, barcode_h = 252, 99
    barcode_img = barcode_img.resize((barcode_w, barcode_h), Image.Resampling.NEAREST)

    x_barcode = (width - barcode_w) // 2
    img.paste(barcode_img, (x_barcode, y_barcode))

    # 4. Alt Numaralar
    sku = data_dict["sku"]
    draw.text(((width - draw.textlength(sku, font=font_sku)) / 2, y_sku), sku, font=font_sku, fill="black")

    margin_side = x_barcode - 50
    dept = data_dict["dept"]
    draw.text((margin_side, y_side), dept, font=font_side, fill="black")

    vendor = data_dict["vendor"]
    sag_bosluk = 30
    draw.text((width - sag_bosluk - draw.textlength(vendor, font=font_side), y_side), vendor, font=font_side,
              fill="black")

    # 5. Kırmızı Kesik Çizgi ve Fiyat
    dash_length = 15
    for x in range(15, width - 15, dash_length * 2):
        draw.line([(x, y_line), (x + dash_length, y_line)], fill="red", width=2)

    if data_dict.get("fiyat"):
        fiyat = data_dict["fiyat"]
        draw.text(((width - draw.textlength(fiyat, font=font_price)) / 2, y_price), fiyat, font=font_price,
                  fill="black")

    # Dosyayı kaydet ve temizle
    img.save(filename)
    if os.path.exists(temp_filename):
        os.remove(temp_filename)

    print(f"Mükemmel uyum! Yeni etiket oluşturuldu: {filename}")


def toplu_etiket_uret(excel_dosyasi):
    # Çıktıların kaydedileceği klasörü oluştur
    output_klasor = "Etiket_Ciktilari"
    if not os.path.exists(output_klasor):
        os.makedirs(output_klasor)

    # Excel dosyasını oku
    print(f"{excel_dosyasi} okunuyor...")
    df = pd.read_excel(excel_dosyasi, dtype={'SKU': str, 'Dept': str, 'Vendor': str})

    # Her satır için döngüye gir
    for index, row in df.iterrows():
        # NaN (boş) hücreleri kontrol et ve string'e çevir
        fiyat_verisi = str(row['Fiyat']) if pd.notna(row['Fiyat']) else ""

        data_dict = {
            "urun_adi": str(row['Urun_Adi']),
            "dept": str(row['Dept']),
            "sku": str(row['SKU']),
            "vendor": str(row['Vendor']),
            "fiyat": fiyat_verisi,
            "brand" : str(row['Brand'])
        }

        # İsimlendirme Mantığı: SKU_fiyatli.png veya SKU_fiyatsiz.png
        if fiyat_verisi.strip():
            # Eğer fiyat alanı doluysa
            dosya_adi = os.path.join(output_klasor, f"{data_dict['sku']}_fiyatli.png")
        else:
            # Eğer fiyat alanı boşsa
            dosya_adi = os.path.join(output_klasor, f"{data_dict['sku']}_fiyatsiz.png")

        # Etiketi oluştur
        create_professional_sticker(data_dict, filename=dosya_adi)
        print(f"Başarılı: {dosya_adi} oluşturuldu.")

    print(f"\nTüm etiketler '{output_klasor}' klasörüne kaydedildi!")


# Script çalıştığında excel dosyasını okuması için tetikliyoruz:
# Not: Excel dosyanızın adı 'urun_listesi.xlsx' ve scriptinizle aynı klasörde olmalı.
if __name__ == "__main__":
    toplu_etiket_uret("sku_sticker_list.xlsx")