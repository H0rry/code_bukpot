import os

from pdf2image import convert_from_bytes, convert_from_path
from pdf2image.exceptions import (
    PDFInfoNotInstalledError,
    PDFPageCountError,
    PDFSyntaxError,
)

print("Hello")


def jalan():
    files = []

    for file in os.listdir("bukpot"):
        if file.endswith(".pdf"):
            files.append(os.path.join(file))

    print(files)

    for file in files:
        filename = file.split(".")[0]

        imgs = convert_from_path(
            f"bukpot/{file}", poppler_path="C:/Program Files/poppler-0.68.0/bin"
        )
        for item in enumerate(imgs):
            item[1].save(f"results/{filename}_page_{item[0]+1}.png")


jalan()
