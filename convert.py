# See https://stackoverflow.com/questions/53481088/poppler-in-path-for-pdf2image and https://towardsdatascience.com/poppler-on-windows-179af0e50150
# conda install -c conda-forge poppler
#    OR
# Install chocolatey https://chocolatey.org/install
# choco install poppler

# python -m pip install --upgrade pip
# python -m pip install --upgrade argparse
# python -m pip install --upgrade pdf2image
# python -m pip install --upgrade Pillow

import argparse
import fnmatch
import os.path
import pathlib
import random
import shutil
from os import path

from pdf2image import convert_from_path
from PIL import Image

# GLOBAL VARIABLES
# ----------------
# Default image size (height of the square).
default_size = 500
languages = {
    "aeat": "Afrikaans",
    "aat": "Afrikaans",
    "aht": "Afrikaans",
    "efal": "English",
    "ehl": "English",
    "ahteat": "Afrikaans",
    "ebw": "Afrikaans",
    "ems": "English",
    "afr": "Afrikaans",
    "eng": "English",
    "pswa": "Afrikaans",
    "pswe": "English",
    "ma": "Afrikaans",
    "me": "English",
    "ns": "English",
    "nst": "English",
    "nw": "Afrikaans",
    "nwt": "Afrikaans",
    "sga": "Afrikaans",
    "sge": "English",
    "sha": "Afrikaans",
    "she": "English",
    "sse": "English",
    "swa": "Afrikaans",
    "teg": "Afrikaans",
    "tswfal": "Setswana",
    "tswhl": "Setswana",
    "tech": "English",
    "vka": "Afrikaans",
    "vke": "English",
    "da": "Afrikaans",
    "de": "English",
    "amus": "Afrikaans",
    "emus": "English",
    "adra": "Afrikaans",
    "edra": "English",
    "eddm": "English",
    "addm": "Afrikaans",
}
subjects = {
    "aeat": "Afrikaans EAT",
    "vka": "Visuele Kuns",
    "vke": "Visual Arts",
    "da": "Dans",
    "de": "Dance",
    "amus": "Musiek",
    "emus": "Music",
    "adra": "Drama (Afrikaans)",
    "edra": "Drama (English)",
    "addm": "Dans Drama Musiek (DDM)",
    "eddm": "Dance Drama Music (DDM)",
    "aat": "Afrikaans AAT",
    "aht": "Afrikaans Huistaal",
    "efal": "English FAL",
    "ehl": "English Home Language",
    "ahteat": "Afrikaans EAT en HT",
    "ebw": "Ekonomiese- en Bestuurswetenskap",
    "ems": "Economic and Management Sciences",
    "afr": "Afrikaans",
    "eng": "English",
    "pswa": "Lewensvaardighede (PSW)",
    "pswe": "Life Skills (PSW)",
    "ma": "Wiskunde",
    "me": "Mathematics",
    "ns": "Natural Sciences",
    "nst": "Natural Sciences and Technology",
    "nw": "Natuurwetenskap",
    "nwt": "Natuurwetenskap en Tegnologie",
    "sga": "Sosiale Wetenskap Aardrykskunde",
    "sge": "Social Sciences Geography",
    "sha": "Sosiale Wetenskap Geskiedenis",
    "she": "Social Sciences History",
    "sse": "Social Sciences",
    "swa": "Sosiale Wetenskap",
    "teg": "Tegnologie",
    "tech": "Technology",
    "tswfal": "Setswana FAL",
    "tswhl": "Setswana Home Language",
}
grades_afr = {
    "gr1": "Gr. 1",
    "gr2": "Gr. 2",
    "gr3": "Gr. 3",
    "gr4": "Gr. 4",
    "gr49": "Gr. 4-9",
    "gr5": "Gr. 5",
    "gr56": "Gr. 5-6",
    "gr57": "Gr. 5-7",
    "gr6": "Gr. 6",
    "gr67": "Gr. 6-7",
    "gr69": "Gr. 6-9",
    "gr7": "Gr. 7",
    "gr79": "Gr. 7-9",
    "gr8": "Gr. 8",
    "gr9": "Gr. 9",
    "gr10": "Gr. 10",
    "gr1012": "Gr. 10-12",
    "gr11": "Gr. 11",
    "gr12": "Gr. 12",
    "junior": "Junior",
    "senior": "Senior",
}
grades_eng = {
    "gr1": "Gr. 1",
    "gr2": "Gr. 2",
    "gr3": "Gr. 3",
    "gr4": "Gr. 4",
    "gr49": "Gr. 4-9",
    "gr5": "Gr. 5",
    "gr56": "Gr. 5-6",
    "gr57": "Gr. 5-7",
    "gr6": "Gr. 6",
    "gr67": "Gr. 6-7",
    "gr69": "Gr. 6-9",
    "gr7": "Gr. 7",
    "gr79": "Gr. 7-9",
    "gr8": "Gr. 8",
    "gr9": "Gr. 9",
    "gr10": "Gr. 10",
    "gr1012": "Gr. 10-12",
    "gr11": "Gr. 11",
    "gr12": "Gr. 12",
    "junior": "Junior",
    "senior": "Senior",
}
grades_list = {
    "gr1": "Gr. 1",
    "gr2": "Gr. 2",
    "gr3": "Gr. 3",
    "gr4": "Gr. 4",
    "gr49": "Gr. 4, Gr. 5, Gr. 6, Gr. 7, Gr. 8, Gr. 9",
    "gr5": "Gr. 5",
    "gr56": "Gr. 5, Gr. 6",
    "gr57": "Gr. 5, Gr. 6, Gr. 7",
    "gr6": "Gr. 6",
    "gr67": "Gr. 6, Gr. 7",
    "gr69": "Gr. 6, Gr. 7, Gr. 8, Gr. 9",
    "gr7": "Gr. 7",
    "gr79": "Gr. 7, Gr. 8, Gr. 9",
    "gr8": "Gr. 8",
    "gr9": "Gr. 9",
    "gr10": "Gr. 10",
    "gr11": "Gr. 11",
    "gr12": "Gr. 12",
    "junior": "Gr. 4, Gr. 5, Gr. 6",
    "senior": "Gr. 7, Gr. 8",
    "": "Gr. 4, Gr. 5, Gr. 6, Gr. 7, Gr. 8, Gr. 9, Gr. 10, Gr. 11, Gr. 12",
}
terms_afr = {
    "k1": "Kwartaal 1",
    "k2": "Kwartaal 2",
    "k3": "Kwartaal 3",
    "k4": "Kwartaal 4",
    "t1": "Kwartaal 1",
    "t2": "Kwartaal 2",
    "t3": "Kwartaal 3",
    "t4": "Kwartaal 4",
    "kw1": "Kwartaal 1",
    "kw2": "Kwartaal 2",
    "kw3": "Kwartaal 3",
    "kw4": "Kwartaal 4",
}
terms_eng = {
    "q1": "Term 1",
    "q2": "Term 2",
    "q3": "Term 3",
    "q4": "Term 4",
    "t1": "Term 1",
    "t2": "Term 2",
    "t3": "Term 3",
    "t4": "Term 4",
}
types_afr = {
    "ass": "Assesseringstaak",
    "bla": "Basislynassesering",
    "sg": "Studiegids",
    "wb": "Werkboek",
    "her": "Hersiening",
    "ppt": "PowerPoint-aanbieding",
    "kv": "Kortverhaal",
    "vkv": "Versameling van kortverhale",
    "altass": "Alternatiewe assesserings",
}
types_eng = {
    "ass": "Assessment Task",
    "bla": "Baseline assessment",
    "sg": "Study Guide",
    "wb": "Workbook",
    "rev": "Revision",
    "ppt": "PowerPoint Presentation",
    "ss": "Short Story",
    "bss": "Bundle of Short Stories",
    "altass": "Alternative Assessments",
}
categories = {
    "ass": "Assessment tasks",
    "bla": "Baseline assessments",
    "sg": "Workbooks and study guides",
    "wb": "Workbooks and study guides",
    "her": "Revision",
    "ppt": "PowerPoint presentations",
    "kv": "Revision",
    "vkv": "PowerPoint presentations",
    "altass": "Assessment tasks",
    "rev": "Revision",
    "ss": "Revision",
    "bss": "PowerPoint presentations",
}


def get_grades(grade):
    gr = grades_list.get(grade)
    if gr is None:
        gr = "Gr. 4, Gr. 5, Gr. 6, Gr. 7, Gr. 8, Gr. 9, Gr. 10, Gr. 11, Gr. 12"
    return gr


def get_category(type):
    category = categories.get(type)
    if category is None:
        category = "Other"
    return category


def make_square(im, min_size=default_size, fill_color=(255, 255, 255, 0)):
    x, y = im.size

    if y > x:
        size = max(min_size, x, y)
        new_im = Image.new("RGBA", (size, y), fill_color)
        new_im.paste(im, (int((size - x) / 2), 0))
        return new_im
    else:
        size = max(min_size, x, y)
        new_im = Image.new("RGBA", (x, size), fill_color)
        new_im.paste(im, (0, int((size - y) / 2)))
        return new_im


def resize_image(im, size=default_size):
    x, y = im.size

    if y > x:
        hpercent = size / float(y)
        hsize = int((float(x)) * float(hpercent))
        new_im = im.resize((hsize, size), Image.Resampling.LANCZOS)
        return new_im
    else:
        vpercent = size / float(x)
        vsize = int((float(y) * float(vpercent)))
        new_im = im.resize((size, vsize), Image.Resampling.LANCZOS)
        return new_im


def move_files(src: str, dst: str, pattern: str = "*"):
    if not os.path.isdir(dst):
        pathlib.Path(dst).mkdir(parents=True, exist_ok=True)
    for f in fnmatch.filter(os.listdir(src), pattern):
        srcfn = os.path.join(src, f)
        if not os.path.isdir(srcfn):
            shutil.move(srcfn, os.path.join(dst, f))


def get_part(parts, pos):
    if len(parts) > pos:
        return parts[pos]
    else:
        return ""


def create_extra(parts, pos):
    extra = ""
    if len(parts) > pos + 1:
        for i in range(pos + 1, len(parts)):
            extra += get_part(parts, i) + " "
        extra = extra.capitalize()
        extra = extra.strip()
    return extra


def create_name(subject, grade, type, year, term, extra):
    prod = ""
    if subject != "":
        prod = prod + subject
    if grade != "":
        prod = prod + " " + grade
    if type != "":
        prod = prod + " " + type
    if year != "":
        prod = prod + " " + year
    if term != "":
        prod = prod + " " + term
    if extra != "":
        prod = prod + " " + extra.replace("_", " ").capitalize()
    prod = prod.strip()
    return prod


def create_product(basename):
    product = ""
    subject = ""
    grade = ""
    gr = ""
    year = ""
    term = ""
    type = ""
    extra = ""
    language = "English"

    pos = 0
    parts = basename.split("_")
    subject = subjects.get(get_part(parts, pos))
    if subject is None:
        product = create_name("", "", "", "", "", basename)
        subject = ""
    else:
        language = languages.get(get_part(parts, pos))
        if language is None:
            language = ""

        pos = pos + 1
        gr = get_part(parts, pos)
        grade = grades_afr.get(gr) if language == "Afrikaans" else grades_eng.get(gr)
        if grade is None:
            grade = ""
        else:
            pos = pos + 1
        year = get_part(parts, pos)
        if year.isnumeric():
            pos = pos + 1
            term = (
                terms_afr.get(get_part(parts, pos))
                if language == "Afrikaans"
                else terms_eng.get(get_part(parts, pos))
            )
            if term is None:
                term = ""
                type = (
                    types_afr.get(get_part(parts, pos))
                    if language == "Afrikaans"
                    else types_eng.get(get_part(parts, pos))
                )
                if type is None:
                    type = ""
                    pos = pos - 1

                extra = create_extra(parts, pos)
            else:
                pos = pos + 1
                type = (
                    types_afr.get(get_part(parts, pos))
                    if language == "Afrikaans"
                    else types_eng.get(get_part(parts, pos))
                )
                if type is None:
                    type = ""
                    pos = pos - 1

                extra = create_extra(parts, pos)
                product = create_name(subject, grade, type, year, term, extra)
        else:
            year = ""
            type = (
                types_afr.get(get_part(parts, pos))
                if language == "Afrikaans"
                else types_eng.get(parts[pos])
            )
            if type is None:
                type = ""
                pos = pos - 1

            extra = create_extra(parts, pos)

        product = create_name(subject, grade, type, year, term, extra)
        product = product.strip()

    subjectCSV = subject
    if subject == "Afrikaans":
        subjectCSV = "Afrikaans EAT, Afrikaans Huistaal"
    if subject == "English":
        subjectCSV = "English FAL, English Home Language"
    if subject == "Afrikaans EAT en HT":
        subjectCSV = "Afrikaans EAT, Afrikaans Huistaal"

    print(
        '0,"simple, downloadable, virtual",,'
        + product
        + ",1,0,visible,,,,,taxable,,1,,,0,1,,,,,1,,,100,"
        + get_category(type)
        + ',"'
        + subjectCSV
        + "\",,,'-1,'-1,,,,,,,0,,,,Language,"
        + language
        + ',0,1,Grade,"'
        + get_grades(gr)
        + '",0,1,,,,,,,,,,,,,,,,,,,,,,'
        + basename
    )
    return product


def process_pdf(filename, outdirroot, size):
    basename = os.path.basename(filename).rsplit(".", 1)[0]
    pdfdir = os.path.dirname(filename)
    outdir = os.path.join(outdirroot, basename)
    outdirimg = os.path.join(outdir, "resized")
    outdirorig = os.path.join(outdir, "orig")
    product = str(create_product(basename)).rstrip()

    if filename.find("memo") != -1:
        return

    try:
        pages = convert_from_path(filename)
    except:
        print("ERROR: Unable to process '" + filename + "'")
        return

    if not path.exists(filename):
        print("ERROR: File does not exist.")
        exit(1)

    if not path.isfile(filename):
        print("ERROR: Supplied filename is not a file.")
        exit(1)

    if not path.exists(outdir):
        os.makedirs(outdir, exist_ok=False)

    if not path.exists(outdirimg):
        os.makedirs(outdirimg, exist_ok=False)

    if not path.exists(outdirorig):
        os.makedirs(outdirorig, exist_ok=False)

    res = []
    nopages = 3
    if len(pages) >= 2:
        res = [2]
    if len(pages) >= 3:
        nopages = min(3, len(pages) - 2)
        res = res + random.sample(range(3, len(pages) + 1), nopages)

    cnt = 0
    for page in pages:
        cnt = cnt + 1

        fnorig = os.path.join(outdirorig, basename + "-" + str(cnt))
        page.save(fnorig + ".png")

        img_resized = resize_image(page, size)

        fnimg = os.path.join(outdirimg, basename + "-" + str(cnt))
        img_resized.save(fnimg + ".png")

        fn = os.path.join(outdir, basename + "-" + str(cnt))

        if cnt in res:
            img_resized.save(fn + ".png")

        if cnt == 1:
            img_new = make_square(img_resized, size)

            fn = os.path.join(outdir, basename + "-preview")
            img_new.save(fn + ".png")

    move_files(pdfdir, outdir, basename + "*")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Export pages from a PDF and convert them to resized and centered images. The files will be created per page. The name will n.png where n starts from 1."
    )
    parser.add_argument(
        "srcdir", metavar="srcdir", type=str, help="The source directory of PDF files."
    )
    parser.add_argument(
        "outdir",
        metavar="outdir",
        type=str,
        help="The output directory for converted files.",
    )
    parser.add_argument(
        "size",
        metavar="size",
        type=int,
        help="Size of the images (height/width of the square).",
        nargs="?",
        default=default_size,
    )
    args = parser.parse_args()

    size = args.size

    if not path.exists(args.outdir):
        os.makedirs(args.outdir, exist_ok=False)

    print(
        'ID,Type,SKU,Name,Published,"Is featured?","Visibility in catalogue","Short description",Description,"Date sale price starts","Date sale price ends","Tax status","Tax class","In stock?",Stock,"Low stock amount","Backorders allowed?","Sold individually?","Weight (kg)","Length (cm)","Width (cm)","Height (cm)","Allow customer reviews?","Purchase note","Sale price","Regular price",Categories,Tags,"Shipping class",Images,"Download limit","Download expiry days",Parent,"Grouped products",Upsells,Cross-sells,"External URL","Button text",Position,"Download 1 ID","Download 1 name","Download 1 URL","Attribute 1 name","Attribute 1 value(s)","Attribute 1 visible","Attribute 1 global","Attribute 2 name","Attribute 2 value(s)","Attribute 2 visible","Attribute 2 global","Download 2 ID","Download 2 name","Download 2 URL","Download 3 ID","Download 3 name","Download 3 URL","Download 4 ID","Download 4 name","Download 4 URL","Download 5 ID","Download 5 name","Download 5 URL","Download 6 ID","Download 6 name","Download 6 URL","Download 7 ID","Download 7 name","Download 7 URL","Download 8 ID","Download 8 name","Download 8 URL",Basename'
    )

    for f in fnmatch.filter(os.listdir(args.srcdir), "*.pdf"):
        srcfn = os.path.join(args.srcdir, f)
        if not os.path.isdir(srcfn):
            process_pdf(srcfn, args.outdir, size)
