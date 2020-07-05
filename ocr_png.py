from PIL import Image
import pytesseract
import re

# (x, y, width, height)
slice_loc = {
    "Name": (248, 1230, 655, 65),
    "NPWP": (250, 1175, 600, 53),
    "No_Bukpot": (680, 390, 450, 50),
    "Date": (945, 1240, 310, 49),
    "Bruto": (580, 1050, 200, 65),
    "PPh_21": (1347, 1050, 238, 63),
}

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract"

lang_list = ["eng", "ind"]


def get_value(file):

    img = Image.open(file)
    results = {}

    for lang in lang_list:
        r_data = {}
        for key, value in slice_loc.items():

            bbox = (
                value[0],
                value[1],
                value[0] + value[2],
                value[1] + value[3],
            )

            chunks_pic = img.crop(bbox)

            ocr = pytesseract.image_to_string(chunks_pic, lang=lang)
            r_data[key] = ocr

            if key != "Name":
                r_data[key] = re.sub("[^0-9]", "", r_data[key])

            if key == "Date":
                r_data[key] = f"{r_data[key][:2]}/{r_data[key][2:4]}/{r_data[key][4:8]}"

            if key == "NPWP":
                r_data[
                    key
                ] = f"{r_data[key][0:2]}.{r_data[key][2:5]}.{r_data[key][5:8]}.{r_data[key][8]}-{r_data[key][9:12]}.{r_data[key][12:]}"

            if key == "No_Bukpot":
                if len(r_data[key]) < 13:
                    r_data[key] = f"1{r_data[key]}"

                if r_data[key][0] != "1":
                    r_data[key] = f"1{r_data[key][1:]}"

                r_data[
                    key
                ] = f"{r_data[key][0]}.{r_data[key][1]}-{r_data[key][2:4]}.{r_data[key][4:6]}-{r_data[key][6:]}"

        r_data.update({"DPP": f"{round(int(r_data['Bruto'])/2)}"})

        sorted_data = {}
        for key in slice_loc.keys():
            if key == "Name":
                sorted_data.update({"File_Pdf": ""})
            if key == "Bruto":
                sorted_data.update({"JENIS PAJAK": "PPh Pasal 21"})
            if key == "PPh_21":
                sorted_data.update({"DPP": r_data["DPP"]})
                # sorted_data.update({'Tarif': '5%'})
            sorted_data.update({key: r_data[key]})

        results[lang] = sorted_data

    return results
