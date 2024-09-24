from PIL import Image
import os
import shutil

sez_dat = os.listdir()

for dat in sez_dat:
    if '.tiff' in dat:
        with Image.open(dat) as slika:
            slika = slika.convert("RGB")
            slika.save(dat.replace('.tiff','') + '.pdf', "PDF")
            os.remove(dat)

posodobljen_sez_dat = os.listdir()
for dat in posodobljen_sez_dat:
    trenutno = os.getcwd()
    premakniti = r"C:\\FileUploader\\WATCHER\\Working"
    if '.pdf' in dat:
        shutil.move(trenutno + '\\' + dat, premakniti + '\\' + dat)