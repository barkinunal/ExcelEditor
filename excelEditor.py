from xlrd import open_workbook
from xlwt import Workbook

class Arm(object):

    howMany = 0

    def __init__(self, hasta_no, protokol_no, adi, soyadi, dogum_tarihi, gelis_tarihi, yasi, poliklik_kodu, servis_adi, kod, tani_adi,lol):
        self.id = id
        self.hasta_no = hasta_no
        self.protokol_no = protokol_no
        self.adi = adi
        self.soyadi = soyadi
        self.dogum_tarihi = dogum_tarihi
        self.gelis_tarihi = gelis_tarihi
        self.yasi = yasi
        self.poliklik_kodu = poliklik_kodu
        self.servis_adi = servis_adi
        self.kod = kod
        self.tani_adi = tani_adi
        self.lol = lol

    def __str__(self):
        return("Arm object:\n"
               "  Hasta no = {0}\n"
               "  Protokol no = {1}\n"
               "  Gelis Tarihi = {2}\n"
               "  Adi = {3}\n"
               "  Soyadi = {4} \n"
               "  Yasi = {5} \n"
               "  Poliklinik kodu = {6} \n"
               "  Servis adi = {7} \n"
               "  Kod = {8} \n"
               "  Tani adi = {9}"

                       .format(self.hasta_no, self.protokol_no, self.gelis_tarihi, self.adi, self.soyadi, self.dogum_tarihi, self.yasi, self.poliklik_kodu, self.servis_adi, self.kod, self.tani_adi))


wb = open_workbook("/Users/yigitbarkinunal/Downloads/L50 POL.xls", encoding_override="cp1254")
yazilacak = open("/Users/yigitbarkinunal/Downloads/hastakayit.txt", "w", encoding='utf-8')
wf = Workbook()
sheet1 = wf.add_sheet('Sheet 1')

for sheet in wb.sheets():
    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols

    items = []
    nolar = []
    rows = []

    for row in range(1, number_of_rows):
        values = []
        for col in range(number_of_columns):
            value  = (sheet.cell(row,col).value)
            try:
                value = str(int(value))
            except ValueError:
                pass
            finally:
                values.append(value)
        item = Arm(*values)
        try:
            if item.gelis_tarihi != '' and int(item.gelis_tarihi[6:]) > 2013:
                if item.hasta_no not in nolar:
                    items.append(item)
                    nolar.append(item.hasta_no)
                else:
                    item.howMany = item.howMany + 1
        except:
            pass


items.sort(key = howMany)
for item in items:
    if item.howMany <= 2:
        item.remove(item)



sheet1.write(0,0,"HASTA_NO")
sheet1.write(0,1,"PROTOKOL_NO")
sheet1.write(0,2,"POLIKLINIK_GELIS_TARIHI")
sheet1.write(0,3,"ADI")
sheet1.write(0,4,"SOYADI")
sheet1.write(0,5,"DOGUM_TARIHI")
sheet1.write(0,6,"YASI")
sheet1.write(0,7,"POLIKLINIK_KODU")
sheet1.write(0,8,"SERVIS_ADI")
sheet1.write(0,9,"KOD")
sheet1.write(0,10,"TANI_ADI")


for j in range(len(items)):
    print(items[j])
    print()

    sheet1.write(j+1,0,items[j].hasta_no)
    sheet1.write(j+1,1,items[j].protokol_no)
    sheet1.write(j+1,2,items[j].gelis_tarihi)
    sheet1.write(j+1,3,items[j].adi)
    sheet1.write(j+1,4,items[j].soyadi)
    sheet1.write(j+1,5,items[j].dogum_tarihi)
    sheet1.write(j+1,6,items[j].yasi)
    sheet1.write(j+1,7,items[j].poliklik_kodu)
    sheet1.write(j+1,8,items[j].servis_adi)
    sheet1.write(j+1,9,items[j].kod)
    sheet1.write(j+1,10,items[j].tani_adi)

    yazilacak.write("{} \n".format(item.hasta_no))
    yazilacak.write("{} \n".format(item.protokol_no))
    yazilacak.write("{} \n".format(item.gelis_tarihi))
    yazilacak.write("{} \n".format(item.adi))
    yazilacak.write("{} \n".format(item.soyadi))
    yazilacak.write("{} \n".format(item.dogum_tarihi))
    yazilacak.write("{} \n".format(item.yasi))
    yazilacak.write("{} \n".format(item.poliklik_kodu))
    yazilacak.write("{} \n".format(item.servis_adi))
    yazilacak.write("{} \n".format(item.kod))
    yazilacak.write("{} \n".format(item.tani_adi))

yazilacak.close()
wf.save("/Users/yigitbarkinunal/Downloads/xlwt example.xlsx")
