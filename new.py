#!/usr/bin/env python3
import openpyxl
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
import time
import os
import easygui


begin = easygui.enterbox("Mulai dari baris:")
to = easygui.enterbox("Hingga baris:")
data = easygui.fileopenbox("Pilih File Excel:")

rows = []
dataframe = openpyxl.load_workbook(data)
dataframe1 = dataframe.active
for i in dataframe1.iter_rows(
        min_row=int(begin)+2, min_col=0, max_row=int(to)+2, max_col=71, values_only=True
):
    rows.append(i)

driver = webdriver.Chrome()
driver.get("http://sensus-kkp.argocipta.com/login")
# driver.find_element(By.NAME,"username").send_keys(os.environ.get("username"))
# driver.find_element(By.ID,"password-field").send_keys(os.environ.get("password"))
# driver.find_element(By.CLASS_NAME,"btn").click()
easygui.msgbox("Login dulu, terus klik OK")

def get_sumber_modal(name):
    name =  name.lower()
    if name== "sendiri" or name == "mandiri":
        return  "sendiri"
    return "kredit"

def get_tanggal_lahir(NIK, gender):
    val = NIK[6:12]
    gender = gender.lower()
    date = int(val[0:2])
    month = int(val[2:4])
    year = int(val[4:6])

    if gender[0] == "p":
        date -= 40

    if year > 22:
        year += 1900
    else:
        year += 2000

    return str(year) + "-" + str(month) + "-" + str(date)

def get_pendidikan(name):
    name = name.lower()
    if name == "s3":
        return 1
    elif name == "s1":
        return 2
    elif name == "s2":
        return 3
    elif name == "d4":
        return 5
    elif name == "d3":
        return 6
    elif name == "d2":
        return 7
    elif name == "d1":
        return 8
    elif name == "sma":
        return 9
    elif name == "smp":
        return 10
    elif name == "sd":
        return 11
    return 99

def select_from_option(value, possible_values):
    for i in range(len(possible_values)):
        if value == possible_values[i]:
            return i
    return len(possible_values) - 1

for i in range(len(rows)):
    num = rows[i][0]
    
    kelompok = rows[i][16]
    biota = rows[i][17]
    komoditas = rows[i][15]
    NIK = rows[i][2]
    gender = rows[i][5]
    agama = rows[i][4]
    nama = rows[i][1]
    tanggal_lahir = get_tanggal_lahir(NIK, gender)
    pendidikan = rows[i][6]
    anggota_keluarga = rows[i][7]
    alamat = rows[i][8]
    desa = rows[i][9]
    kecamatan = rows[i][10]

    longitude = rows[i][14]
    latitude = rows[i][13]

    jenis_usaha = rows[i][18]
    status_kusuka = rows[i][19]

    kepemilikan = rows[i][20]
    luas = rows[i][21]
    media = rows[i][22]
    tekonologi = rows[i][24]

    produktifitas = rows[i][27]
    harga_jual = rows[i][29]
    pendapatan = rows[i][30]

    jenis_pakan = rows[i][31]
    jumlah_pakan = rows[i][32]
    harga_pakan = rows[i][36]
    harga_pembelian_pakan = rows[i][37]

    jumlah_benih = rows[i][40]
    harga_benih = rows[i][41]
    harga_pembelian = rows[i][39]

    jumlah_tk = rows[i][43]
    besaran_modal = rows[i][44]
    sumber_modal = get_sumber_modal(rows[i][45])
    biaya_pembuatan_media = rows[i][47]
    biaya_penyusutan = rows[i][48]
    biaya_peralatan = rows[i][49]
    biaya_tenaga_kerja = rows[i][50]

    IPAL = rows[i][51]
    tandon = rows[i][52]

    green_belt = rows[i][53]
    jarak_ke_pantai = rows[i][54]
    sumber_air = rows[i][55]

    perizinan = rows[i][56]
    status_NIB = rows[i][57]
    skala_usaha = rows[i][58]

    asuransi = rows[i][59]
    bantuan = rows[i][60]
    penghargaan = rows[i][61]
    dukungan_pemda = rows[i][62]
    dukungan_pusat = rows[i][63]
    sertifikat = rows[i][65]
    nama_penyuluh = rows[i][64]

    # print(
    #     kelompok,
    #     biota,
    #     komoditas,
    #     NIK,
    #     tanggal_lahir,
    #     pendidikan,
    #     anggota_keluarga,
    #     alamat,
    #     desa,
    #     kecamatan,
    #     longitude,
    #     latitude,
    #     jenis_usaha,
    #     status_kusuka,
    #     kepemilikan,
    #     luas,
    #     media,
    #     tekonologi,
    #     produktifitas,
    #     harga_jual,
    #     pendapatan,
    #     jenis_pakan,
    #     jumlah_pakan,
    #     harga_pakan,
    #     harga_pembelian_pakan,
    #     jumlah_benih,
    #     harga_benih,
    #     harga_pembelian,
    #     jumlah_tk,
    #     besaran_modal,
    #     sumber_modal,
    #     biaya_pembuatan_media,
    #     biaya_penyusutan,
    #     biaya_peralatan,
    #     biaya_tenaga_kerja,
    #     IPAL,
    #     tandon,
    #     green_belt,
    #     jarak_ke_pantai,
    #     sumber_air,
    #     perizinan,
    #     status_NIB,
    #     asuransi,
    #     bantuan,
    #     penghargaan,
    #     dukungan_pemda,
    #     dukungan_pusat,
    #     sertifikat,
    #     nama_penyuluh)

    driver.get("http://sensus-kkp.argocipta.com/admin/sensus/rtp/new")

    Select(driver.find_element(By.NAME, "kelompok_id")).select_by_visible_text(kelompok)
    Select(driver.find_element(By.NAME, "biota_id")).select_by_visible_text(biota)
    time.sleep(2)
    Select(driver.find_element(By.NAME, "ikan_id")).select_by_visible_text(komoditas.title())

    driver.find_element(By.NAME, "nik").send_keys(NIK)
    Select(driver.find_element(By.NAME, "agama_id")).select_by_visible_text(agama.title())
    driver.find_element(By.NAME, "name").send_keys(nama)
    driver.find_element(By.NAME, "birthdate").send_keys(tanggal_lahir)
    Select(driver.find_element(By.NAME, "pendidikan")).select_by_value(str(get_pendidikan(pendidikan)))
    driver.find_element(By.NAME, "family_num").send_keys(anggota_keluarga)
    
    driver.find_element(By.NAME, "address").send_keys(alamat)
    Select(driver.find_element(By.NAME, "kecamatan_id")).select_by_visible_text(kecamatan)

    Select(driver.find_element(By.NAME, "kelurahan_id")).select_by_visible_text(desa)
    driver.find_element(By.NAME, "lat").send_keys(latitude)
    driver.find_element(By.NAME, "lng").send_keys(longitude)

    Select(driver.find_element(By.NAME, "jenis_usaha")).select_by_visible_text(jenis_usaha.title())
    Select(driver.find_element(By.NAME, "status_kusuka")).select_by_visible_text(status_kusuka)
    Select(driver.find_element(By.NAME, "status_milik")).select_by_visible_text(kepemilikan)
    driver.find_element(By.NAME, "luas_usaha").send_keys(luas)
    driver.find_element(By.NAME, "media_pelihara").send_keys(media)
    Select(driver.find_element(By.NAME, "teknologi")).select_by_visible_text(tekonologi)


    driver.find_element(By.NAME, "produksi").send_keys(produktifitas)
    driver.find_element(By.NAME, "harga").send_keys(harga_jual)
    driver.find_element(By.NAME, "income").send_keys(pendapatan)

    Select(driver.find_element(By.NAME, "pakan_jenis")).select_by_visible_text(jenis_pakan.title())
    driver.find_element(By.NAME, "pakan_num").send_keys(jumlah_pakan)
    driver.find_element(By.NAME, "pakan_harga").send_keys(harga_pakan)
    driver.find_element(By.NAME, "biaya_pakan").send_keys(harga_pembelian_pakan)

    driver.find_element(By.NAME, "benur_num").send_keys(jumlah_benih)
    driver.find_element(By.NAME, "benur_harga").send_keys(harga_benih)
    driver.find_element(By.NAME, "biaya_benih").send_keys(harga_pembelian)
    
    driver.find_element(By.NAME, "tk_num").send_keys(jumlah_tk)
    driver.find_element(By.NAME, "omzet").send_keys(besaran_modal)

    Select(driver.find_element(By.NAME, "sumber_modal")).select_by_visible_text(sumber_modal.title())
    driver.find_element(By.NAME, "biaya_media").send_keys(biaya_pembuatan_media)
    driver.find_element(By.NAME, "biaya_susut").send_keys(biaya_penyusutan)
    driver.find_element(By.NAME, "biaya_alat").send_keys(biaya_peralatan)
    driver.find_element(By.NAME, "biaya_tk").send_keys(biaya_tenaga_kerja)

    Select(driver.find_element(By.NAME, "ipal")).select_by_visible_text(IPAL.title())
    Select(driver.find_element(By.NAME, "tandon")).select_by_visible_text(tandon.title())
    Select(driver.find_element(By.NAME, "greenbelt")).select_by_visible_text(green_belt.title())
    driver.find_element(By.NAME, "jarak_tambak").send_keys(jarak_ke_pantai)
    driver.find_element(By.NAME, "sumber_air").send_keys(sumber_air)

    Select(driver.find_element(By.NAME, "izin")).select_by_visible_text(perizinan.title())
    Select(driver.find_element(By.NAME, "nib")).select_by_visible_text(status_NIB.title())
    Select(driver.find_element(By.NAME, "skala_usaha")).select_by_visible_text(skala_usaha.title())

    Select(driver.find_element(By.NAME, "asuransi")).select_by_visible_text(asuransi.title())
    Select(driver.find_element(By.NAME, "bantuan")).select_by_visible_text(bantuan.title())
    Select(driver.find_element(By.NAME, "penghargaan")).select_by_visible_text(penghargaan.title())
    Select(driver.find_element(By.NAME, "dukungan_pemda")).select_by_visible_text(dukungan_pemda.title())
    Select(driver.find_element(By.NAME, "dukungan_pusat")).select_by_visible_text(dukungan_pusat.title())
    Select(driver.find_element(By.NAME, "sertifikat")).select_by_visible_text(sertifikat.title())

    driver.find_element(By.NAME, "penyuluh_name").send_keys(nama_penyuluh)

    easygui.msgbox("Check dulu datanya lengkap atau tidak. Kalau sudah yakin, klik OK untuk lanjut")
    driver.find_element(By.ID, "submitUser").click()

    time.sleep(4)



print("Done")
