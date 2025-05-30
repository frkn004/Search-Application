import customtkinter as ctk
import autoSel as automation
from threading import Thread
import toExcel as excel
import requests

# Fonksiyonlar

def get_api():
    api_url = "https://turkiyeapi.dev/api/v1/provinces"
    response = requests.get(api_url)
    response.json()
    print(response)

def optionmenu_callbackCity(choice):
    global searchCity
    searchCity = ""
    searchCity = choice
    return str(choice)

def optionmenu_callback(choice):
    optionmenuUnderCity.configure(values=getUnderCityNameData(str(choice)))
    global searchUnderCity
    searchUnderCity = ""
    searchUnderCity = choice

def getCityNameData():
    global cityNameData
    cityNameData = []
    api_url = "https://turkiyeapi.dev/api/v1/provinces"
    response = requests.get(api_url)
    cityData = response.json()
    for i in range(81):
        cityNameData.append(str(cityData["data"][i]["name"]))
    return cityNameData

def getUnderCityNameData(cityName):
    global underCityNameData
    underCityNameData = []
    api_url = "https://turkiyeapi.dev/api/v1/provinces?name="+cityName
    response = requests.get(api_url)
    underCityData = response.json()
    for city in underCityData["data"]:
       for i in (city["districts"]):
            underCityNameData.append(i["name"])
    return underCityNameData

def get_text(textBox):
    text = textBox.get("1.0", "end-1c") 
    return text

def on_enter(event):
    return "break"

def threadingSearch():
    sendButton.configure(state="disabled")
    t1=Thread(target=go_search)
    t1.start()

def loading_popup(message):
    popup = ctk.CTkToplevel(app)
    popup.geometry("500x70")
    popup.title("Bilgilendirme")
    popup.resizable(False,False)
    popup.attributes('-topmost', True)
    labelInfo = ctk.CTkLabel(popup,text=message, font=("Arial",18))
    labelInfo.pack(pady=13,padx=13)
    return popup

def go_search():
    get_api()
    if(get_text(textBoxKategori)!=""):
        try:
            global popupLoading
            text = get_text(textBoxKategori) + ", " + get_text(textBoxMahalle) + ", " + searchUnderCity + ", " + searchCity
            popupLoading = loading_popup("Veriler Çekiliyor, Lütfen Bekleyiniz...")
            company_data = automation.search(text)
            populate_data(company_data)
        except Exception as e:
            global popupError
            popupError = loading_popup(e)
            sendButton.configure(state="normal")
    else:
        global popupErrorCate
        popupErrorCate = loading_popup("Aranacak Kategori Giriniz.")
        sendButton.configure(state="normal")

def populate_data(data):
    # Scrollable Frame
    scrollable_frame = ctk.CTkScrollableFrame(app,width=1000)
    scrollable_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
    parsed_data = []
    
    print(f"İşlenecek veri sayısı: {len(data)}")
    
    for item in data:
        try:
            # Eğer veri boş ise atla
            if not item or len(item.strip()) == 0:
                continue
                
            # Satırları ayır
            split_data = item.split('\n')
            print(f"Satır sayısı: {len(split_data)}")
            
            # En az 2 satır bilgi olmalı
            if len(split_data) < 2:
                continue
                
            # İşletme adı her zaman ilk satırdır
            company_name = split_data[0]
            
            # Adres, puan ve iletişim bilgilerini bulmaya çalış
            rating = "Belirtilmemiş"
            address = "Belirtilmemiş"
            phone = "Belirtilmemiş"
            
            # Satırlarda telefon numarası, adres ve puan bilgilerini ara
            for line in split_data[1:]:
                line = line.strip()
                if not line:
                    continue
                
                # Telefon numarası kontrolü
                if any(phone_indicator in line for phone_indicator in ['+', '(0', '05', '(5']) or line.replace(' ', '').isdigit():
                    phone = line
                # Puan kontrolü (x.x/5 formatında veya sadece x.x)
                elif '/' in line or ('.' in line and len(line) <= 5):
                    try:
                        # Puan sayı mı kontrol et
                        parts = line.split('/')
                        float(parts[0])  # Sayı mı diye kontrol et
                        rating = line
                    except ValueError:
                        # Sayı değilse, muhtemelen adres
                        if not 'Belirtilmemiş' in address:
                            address = line
                # Diğer durumlar muhtemelen adres
                elif len(line) > 5 and not rating in line and not phone in line:
                    address = line
            
            parsed_data.append({
                'Şirket Adı': company_name,
                'Adres': address + " / " + searchUnderCity + ", " + searchCity,
                'Puan': rating,
                'İletişim': phone
            })
            
            print(f"İşletme eklendi: {company_name}")
            
        except Exception as e:
            print(f"Veri işleme hatası: {e}")
            continue

    print(f"İşlenen toplam işletme sayısı: {len(parsed_data)}")
    
    for widget in scrollable_frame.winfo_children():
        widget.destroy()
    
    if not parsed_data:
        # Eğer işlenen veri yoksa bilgi mesajı göster
        no_data_label = ctk.CTkLabel(scrollable_frame, text="Arama kriterlerine uygun işletme bulunamadı.", font=("Arial", 14))
        no_data_label.grid(row=0, column=0, padx=10, pady=20, sticky="w")
    else:
        # Veri varsa normal işleme devam et
        # Sütun etiketleri
        labels = ['Adı', 'Adres', 'Puan', 'İletişim']
        for i, label_text in enumerate(labels):
            label = ctk.CTkLabel(scrollable_frame, text=label_text, font=("Arial",20))
            label.grid(row=0, column=i, pady=5, padx=5, sticky="w")
            
        for i, company in enumerate(parsed_data, start=1):
            ctk.CTkLabel(scrollable_frame, text=company['Şirket Adı']).grid(row=i, column=0, padx=5, pady=5, sticky="w")
            ctk.CTkLabel(scrollable_frame, text=company['Adres']).grid(row=i, column=1, padx=5, pady=5, sticky="w")
            ctk.CTkLabel(scrollable_frame, text=company['Puan']).grid(row=i, column=2, padx=5, pady=5, sticky="w")
            ctk.CTkLabel(scrollable_frame, text=company['İletişim']).grid(row=i, column=3, padx=5, pady=5, sticky="w")    

    app.geometry("900x700")
    app.grid_rowconfigure(1, weight=1)
    app.grid_columnconfigure(0, weight=1)
    app.grid_rowconfigure(0, weight=1)
    
    print(len(parsed_data))
    excel.get_data(parsed_data)
    
    sendButton.configure(
        state="normal",
        width=100, 
        height=15, 
        text="Gönder", 
        command=threadingSearch
    )
    popupLoading.destroy()
    global popup
    popup = loading_popup("Veriler Çekildi, Excel Dosyası Kaydedildi...")

# Ana uygulama
app = ctk.CTk()
app.title("İşletme Arama")
app.geometry("600x400")
app.minsize(500, 300)
app.grid_columnconfigure(0, weight=1)
app.grid_rowconfigure(0, weight=1)
app.iconbitmap('icon.ico')


# Widgetlar
contFrame = ctk.CTkFrame(app)
contFrame.grid_columnconfigure(0, weight=1)
contFrame.grid_rowconfigure(0, weight=1)
contFrame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

frame = ctk.CTkFrame(contFrame, fg_color="transparent")
frame.grid()

# Kategori TextBox
label = ctk.CTkLabel(frame, text="Aranacak Kategori:")
label.grid(row=0, column=0, padx=10, pady=10)
textBoxKategori = ctk.CTkTextbox(frame, height=25, width= 140)
textBoxKategori.grid(row=0, column=1, padx=10, pady=10)

# Mahalle TextBox
label = ctk.CTkLabel(frame, text="Aranacak Mahalle(Opsiyonel):")
label.grid(row=1, column=0, padx=10, pady=10)
textBoxMahalle = ctk.CTkTextbox(frame, height=25, width= 140)
textBoxMahalle.grid(row=1, column=1, padx=10, pady=10)

# İl OptionMenu
label = ctk.CTkLabel(frame, text="Aranacak İl:")
label.grid(row=2, column=0, padx=10, pady=10)
optionmenu_varCity = ctk.StringVar(value="Seçiniz...")
optionmenuCity = ctk.CTkOptionMenu(frame,values=getCityNameData(),command=optionmenu_callback,variable=optionmenu_varCity)
optionmenuCity.grid(row=2, column=1, padx=10, pady=10)

# İlçe OptionMenu
label = ctk.CTkLabel(frame, text="Aranacak İlçe:")
label.grid(row=3, column=0, padx=10, pady=10)
optionmenu_varUnderCity = ctk.StringVar(value="Seçiniz...")
optionmenuUnderCity = ctk.CTkOptionMenu(frame,values=[""],command=optionmenu_callbackCity,variable=optionmenu_varUnderCity)
optionmenuUnderCity.grid(row=3, column=1, padx=10, pady=10)

# Gönder Butonu
sendButton = ctk.CTkButton(
    frame, 
    width=100, 
    height=5, 
    text="Gönder", 
    command=threadingSearch
)
sendButton.grid(row=4, columnspan=2, ipady=5, ipadx=5, padx=5, pady=5, sticky="nsew")

app.mainloop()
