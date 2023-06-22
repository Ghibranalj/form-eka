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
    min_row=int(begin) + 2, min_col=0, max_row=int(to) + 2, max_col=71, values_only=True
):
    rows.append(i)

driver = webdriver.Chrome()
driver.delete_all_cookies()
driver.get("http://sensus-kkp.argocipta.com/login")
# driver.find_element(By.NAME,"username").send_keys(os.environ.get("username"))
# driver.find_element(By.ID,"password-field").send_keys(os.environ.get("password"))
# driver.find_element(By.CLASS_NAME,"btn").click()
easygui.msgbox("Login dulu, terus klik OK")


def slow_type(element, text, delay=0.0005):
    element.send_keys(" ")
    element.clear()
    element.send_keys(text)


def get_sumber_modal(name):
    name = name.lower()
    if name == "sendiri" or name == "mandiri":
        return "sendiri"
    return "kredit"


def get_tanggal_lahir(NIK, gender):

    val = str(NIK)[6:12]
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
    NIK = str(rows[i][2])
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
    #     nama,
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
    time.sleep(5)
    Select(driver.find_element(By.NAME, "ikan_id")).select_by_visible_text(
        komoditas.title()
    )

    # cahnge all send_keys to slow_type()

    slow_type(driver.find_element(By.NAME, "nik"), NIK)
    time.sleep(1)
    Select(driver.find_element(By.NAME, "agama_id")).select_by_visible_text(
        agama.title()
    )
    slow_type(driver.find_element(By.NAME, "name"), nama, delay=0.03)
    slow_type(driver.find_element(By.NAME, "birthdate"), tanggal_lahir)
    Select(driver.find_element(By.NAME, "pendidikan")).select_by_value(
        str(get_pendidikan(pendidikan))
    )
    slow_type(driver.find_element(By.NAME, "family_num"), str(anggota_keluarga))

    slow_type(driver.find_element(By.NAME, "address"), alamat, delay=0.03)
    Select(driver.find_element(By.NAME, "kecamatan_id")).select_by_visible_text(
        kecamatan
    )

    Select(driver.find_element(By.NAME, "kelurahan_id")).select_by_visible_text(desa)

    slow_type(driver.find_element(By.NAME, "lat"), str(latitude), delay=0.03)

    slow_type(driver.find_element(By.NAME, "lng"), str(longitude), delay=0.03)

    Select(driver.find_element(By.NAME, "jenis_usaha")).select_by_visible_text(
        jenis_usaha.title()
    )
    Select(driver.find_element(By.NAME, "status_kusuka")).select_by_visible_text(
        status_kusuka
    )
    Select(driver.find_element(By.NAME, "status_milik")).select_by_visible_text(
        kepemilikan
    )

    slow_type(driver.find_element(By.NAME, "luas_usaha"), str(luas))

    Select(driver.find_element(By.NAME, "media_pelihara")).select_by_visible_text(
        media.title()
    )

    Select(driver.find_element(By.NAME, "teknologi")).select_by_visible_text(tekonologi)

    slow_type(driver.find_element(By.NAME, "produksi"), str(produktifitas))
    slow_type(driver.find_element(By.NAME, "harga"), str(harga_jual))
    slow_type(driver.find_element(By.NAME, "income"), str(pendapatan))

    Select(driver.find_element(By.NAME, "pakan_jenis")).select_by_visible_text(
        jenis_pakan.title()
    )
    slow_type(driver.find_element(By.NAME, "pakan_num"), str(jumlah_pakan))
    slow_type(driver.find_element(By.NAME, "pakan_harga"), str(harga_pakan))
    slow_type(driver.find_element(By.NAME, "biaya_pakan"), str(harga_pembelian_pakan))

    slow_type(driver.find_element(By.NAME, "benur_num"), str(jumlah_benih))
    slow_type(driver.find_element(By.NAME, "benur_harga"), str(harga_benih))
    slow_type(driver.find_element(By.NAME, "biaya_benih"), str(harga_pembelian))

    slow_type(driver.find_element(By.NAME, "tk_num"), str(jumlah_tk))
    slow_type(driver.find_element(By.NAME, "omzet"), str(besaran_modal))

    Select(driver.find_element(By.NAME, "sumber_modal")).select_by_visible_text(
        sumber_modal.title()
    )
    slow_type(driver.find_element(By.NAME, "biaya_media"), str(biaya_pembuatan_media))
    slow_type(driver.find_element(By.NAME, "biaya_susut"), str(biaya_penyusutan))
    slow_type(driver.find_element(By.NAME, "biaya_alat"), str(biaya_peralatan))
    slow_type(driver.find_element(By.NAME, "biaya_tk"), str(biaya_tenaga_kerja))

    Select(driver.find_element(By.NAME, "ipal")).select_by_visible_text(IPAL.title())
    Select(driver.find_element(By.NAME, "tandon")).select_by_visible_text(
        tandon.title()
    )
    Select(driver.find_element(By.NAME, "greenbelt")).select_by_visible_text(
        green_belt.title()
    )
    slow_type(driver.find_element(By.NAME, "jarak_tambak"), str(jarak_ke_pantai))
    slow_type(driver.find_element(By.NAME, "sumber_air"), sumber_air)

    Select(driver.find_element(By.NAME, "izin")).select_by_visible_text(
        perizinan.title()
    )
    Select(driver.find_element(By.NAME, "nib")).select_by_visible_text(
        status_NIB.title()
    )
    Select(driver.find_element(By.NAME, "skala_usaha")).select_by_visible_text(
        skala_usaha.title()
    )

    Select(driver.find_element(By.NAME, "asuransi")).select_by_visible_text(
        asuransi.title()
    )
    Select(driver.find_element(By.NAME, "bantuan")).select_by_visible_text(
        bantuan.title()
    )
    Select(driver.find_element(By.NAME, "penghargaan")).select_by_visible_text(
        penghargaan.title()
    )
    Select(driver.find_element(By.NAME, "dukungan_pemda")).select_by_visible_text(
        dukungan_pemda.title()
    )
    Select(driver.find_element(By.NAME, "dukungan_pusat")).select_by_visible_text(
        dukungan_pusat.title()
    )
    Select(driver.find_element(By.NAME, "sertifikat")).select_by_visible_text(
        sertifikat.title()
    )

    slow_type(driver.find_element(By.NAME, "penyuluh_name"), nama_penyuluh)

    easygui.msgbox(
        "Check dulu datanya lengkap atau tidak. Kalau sudah yakin, klik OK untuk lanjut"
    )
    driver.find_element(By.ID, "submitUser").click()

    if not easygui.ynbox("Lanjut?"):
        break


easygui.msgbox("Selesai...")
