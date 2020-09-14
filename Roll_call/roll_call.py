import xlrd
from tkinter import *
from tkinter import filedialog
class Editor(Frame):

    def __init__(self,parent):
        Frame.__init__(self, parent)
        self.root=parent
        self.initUI()

    def initUI(self):
        self.grid()
        frame = Frame(self,bg="Beige" , width = "700", height="500")
        frame.grid()
        # canvas=Canvas(self)
        # canvas.create_line(0,30,1500,30)
        # canvas.grid()
        #### ALT KISIMDA BUTON LABEL ENTRY VE TEXT'LERINI OLUSTURUP YERINIZ AYARLIYORZ-
        #### YERLERIRNI SATIR SUTUN SEKLINDE VE ARKA PLAN RENKLERINI VE YAZI FONTLARINI AYARLIYORUZ..
        self.boslabel = Label(frame,bg="Beige")
        self.boslabel.grid(row = 1,column = 0, columnspan = 2)
        self.boslabel2 = Label(frame, bg="Beige")
        self.boslabel2.grid(row=3, column=0, columnspan=2)
        self.boslabel3 = Label(frame,bg="Beige")
        self.boslabel3.grid(row=5, column=0, columnspan=2)
        self.boslabel4 = Label(frame, bg="Beige")
        self.boslabel4.grid(row=7, column=0, columnspan=2)
        self.baslik = Label(frame,text="YOKLAMA HESAPLAYICI",fg="Blue",bg ="Beige")
        self.baslik.config(font=("Courier", 16,"bold italic"))
        self.baslik.grid(row = 0,column = 0, columnspan = 2)
        self.button=Button(frame,text="Dosya Seç",fg="Blue",bg ="Gold",command=self.Dosya)#Command kısmında isi yapacak olan fonksiyona yolluyoruz emri.
        self.var = StringVar()
        self.Yoklamaoran = Entry(frame, textvariable=self.var,width=40,bg="Light blue",fg = "black")
        self.dosyaSeç=Label(frame,text="Excell dosyası seçin:",fg="Blue",bg ="Beige")
        self.YoklamaGir=Label(frame,text="Yoklama oranını giriniz(%):",fg="Blue",bg ="Beige")
        self.goster = Button(frame,text="Göster",fg="Blue",bg ="Gold" , command=self.Goster)#yapacagi isi goster fonksiyonunda bulmasını soyluyoruz.
        self.temizle=Button(frame,text="Temizle",fg="Blue",bg ="Gold",command=self.Temizle)#emir olarak temizle fonk'a gidip orada ki isi yapmasını soyluyoruz.
        self.text = Text(frame, font="Times 13" ,height=15,bg="Beige")

    
        self.dosyaSeç.grid(row = 2,column = 0)
        self.button.grid(row = 2, column = 1,sticky="W",padx="83")
        self.YoklamaGir.grid(row = 4, column = 0)
        self.Yoklamaoran.grid(row = 4, column = 1)
        self.goster.grid(row = 6, column = 1,sticky="W",padx="83")
        self.temizle.grid(row = 6, column = 1)

        self.text.grid(row = 8,column = 0, columnspan = 2)
    def Dosya(self):
        self.filename = filedialog.askopenfilename(initialdir="/", title="Dosya Seç",
                                                      filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))

        #xlrd değişkenine gömmek için dosya yolunun bulunması

    def Goster(self):
        try:
            self.oran=self.var.get()                
            self.text.delete("1.0",END)
            if int(self.oran)>100 or int(self.oran)<0:
                self.text.insert(INSERT,"LUTFEN EN FAZLA %100, EN AZ %0 ORANINDA DEGER GİRİNİZ.")
            elif int(self.oran) == 0:
                self.text.insert(INSERT,"Devamsızlık oranı 0 oldugu icin derse katılım zorunlu degildir.")   
            else:    
                #dosyaya erişmek için gerekli xlrd değişkenleri
                book = xlrd.open_workbook(self.filename)
                sayfa = book.sheet_by_index(0)
                gelmedigiGun = 0 #gelmediği günleri hesaplama bayrağı
                ilk_section_devamsizlikSayisi=[] #ilk section için gerekli devamsızlık dizisi
                ikinci_section_devamsizlikSayisi =[] #ikinci section için gerekli devamsızlık dizisi
                kalan_ogrenci_sayisi=0 #kalan ogrenci sayilari icin sayac tuttuk
                toplam_ders_sayisi=14 #bir dönemde 14 hafta oldugu icin degisken belirledik
                toplam_kalan_ogrenci_sayisi=0 #toplam kalan ogrencileri yazmak icin bir sayac tuttuk
                #sil dediğim kısımda ders sayısını bulmak için sonra sil/sayfa.nrows yaparsak 14 ü elde ederiz.
                #Satirlarda ve sütunlarda gezdik, devamsızlığın olduğu sütunlarda x işaretlerini okuyup gelinmeyen günler bayrak yardımıyla bulunmuştur.
                for i in range (0,int(sayfa.nrows)):
                    sectionSatiri = str(sayfa.cell(i,1).value)
                    for sutun_no in range(6,sayfa.ncols):
                        a = str(sayfa.cell(i, sutun_no).value)
                        if(a == "O"):
                            gelmedigiGun += 1
                #Sectionlara göre devamsızlık yapan öğrencilerimizi ayrı dizilerde tutuyoruz. Section bilgisini excel dosyasından çekiyoruz
                    if(sectionSatiri[0] == '1'):
                        ilk_section_devamsizlikSayisi.append(14-gelmedigiGun)
                    elif(sectionSatiri[0] == '2'):
                        ikinci_section_devamsizlikSayisi.append(14-gelmedigiGun)
                    gelmedigiGun = 0
                #Bu bölümde kullanıcıdan girilen yüzde oranına göre ilk_section_devamsizlikSayisi adlı dizinimizdeki öğrencilerin devamsızlık sayısına göre yüzdelik oranı karşılaştırıp
                #Kalıp kalmadığına karar veriyoruz. Bir bayrak yardımıyla toplam kalan öğrenci sayısını elde ediyoruz.
                #Toplam değişkeni ise tüm sectionlarda kalan öğrenci sayısını bize veriyor.
                sınır=toplam_ders_sayisi*(int(self.oran)/100)  #kaç gün gelmez ise sıkıntı cıkmayacak ise o gun sayisini bulduk
                for i in range(len(ilk_section_devamsizlikSayisi)):
                    if ilk_section_devamsizlikSayisi[i] <= sınır and ilk_section_devamsizlikSayisi[i] != 14:
                        kalan_ogrenci_sayisi+=1
                toplam_kalan_ogrenci_sayisi+=kalan_ogrenci_sayisi
                self.text.insert(INSERT,"1. Sectiondan kalan ogrenci sayisi: "+str(kalan_ogrenci_sayisi)+"\n\n\n")
                kalan_ogrenci_sayisi=0
                for j in range(len(ikinci_section_devamsizlikSayisi)):
                    if ikinci_section_devamsizlikSayisi[j] <= sınır and ikinci_section_devamsizlikSayisi[j] != 14:
                        kalan_ogrenci_sayisi+=1
                toplam_kalan_ogrenci_sayisi+= kalan_ogrenci_sayisi
                self.text.insert(INSERT,"2.sectiondan kalan ogrenci sayisi: "+str(kalan_ogrenci_sayisi)+"\n\n\n")
                self.text.insert(INSERT,"Toplam kalan ogrenci sayisi: "+str(toplam_kalan_ogrenci_sayisi)+"\n\n")
        #Dosya seçmeme, Value Error ve genel errorlar icin hata ayıklması yapıyoruz.
        except AttributeError:
            self.text.insert(INSERT,"Dosya seçiniz!!!")
        except FileNotFoundError:
            self.text.insert(INSERT,"Dosya seçiniz!!!")
        except ValueError:
            self.text.insert(INSERT,"Devamsızlık oranı giriniz!!!")
        except:
            self.text.insert(INSERT,"Beklenmedik bir hata oluştu..")
    def Temizle(self):
        self.text.delete("1.0", END)
        self.Yoklamaoran.delete(first=0, last=100)
        self.filename = ""
        #Temizle butonuna bastıgımızda girilen degerleri temizliyoruz.

def main():
    root= Tk()
    root.title("Excel-Reader")
    root.geometry("700x500+300+100")
    #konumu ayarlıyoruz ve ekran boyut ayarlamasını kapatıyoruz.
    root.resizable(FALSE,FALSE)
    App = Editor(root)
    root.mainloop()

if __name__ == '__main__':
    main()