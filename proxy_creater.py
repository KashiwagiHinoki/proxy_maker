import pandas as pd
from mtgsdk import Card
import ssl
import urllib.request
from PIL import Image
import re
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from tqdm import tqdm

def read_deck_list(file_name):
    df = pd.read_excel(file_name, header=None)
    return df

def split_card_name(card_name):
    __card_name = __card_name = re.sub("\s*(\+|//)\s*", "\n", card_name)
    _card_name = re.split("\n", __card_name)
    return _card_name

def is_split_card(card_name):
    if(len(split_card_name(card_name)) >= 2):
        return True
    else:
        False

def get_card_image(card_name):
    ssl._create_default_https_context = ssl._create_unverified_context
    spliting = is_split_card(card_name)
    if spliting:
        card_name = split_card_name(card_name)[0]    

    card_list = Card.where(language="Japanese", name=card_name).all()
    if(len(card_list) <= 0):
        print("cannot search card name: " + card_name)
        return

    for i in range(len(card_list)):
        for j in range(len(card_list[i].foreign_names)):
            foreign_card = card_list[i].foreign_names[j]
            if foreign_card['language'] == 'Japanese' and foreign_card['name'] == card_name and foreign_card['imageUrl'] != None:
                image_url = foreign_card['imageUrl']

    r = urllib.request.urlopen(image_url)
    card_image_bin = r.read()
    r.close()
    card_image = Image.open(io.BytesIO(card_image_bin))
    return card_image

def get_pdf_image(pdf_name, card_image_list):
    cnvs = canvas.Canvas(pdf_name)
    cnvs.setPageSize(A4)
    rows = 3
    columns = 3
    x_card_size = 63*0.98*mm
    y_card_size = 88*0.98*mm
    x_margin = (A4[0] - 3 * x_card_size) / 4
    y_margin = (A4[1] - 3 * y_card_size) / 4

    for i in range(len(card_image_list)):
        row = (i % (rows * columns)) // rows
        column = i % columns
        if i != 0 and row == 0 and column == 0:
            cnvs.showPage()
        
        y = row * (y_margin + y_card_size) + y_margin
        x = column * (x_margin + x_card_size) + x_margin
        cnvs.drawInlineImage(card_image_list[i], x, y, x_card_size, y_card_size)

    cnvs.save()
    return

if __name__ == '__main__':
    print("input excel format file name > ", end="")
    file_name = input()
    print("input created pdf format file name > ", end="")
    output_file_name = input()
    deck_list = read_deck_list(file_name)
    
    cards_list = deck_list.iloc[:, 0]
    card_image_list = []
    for card_name in tqdm(cards_list):
        card_image_list.append(get_card_image(card_name))

    deck_image_list = []
    for i in range(len(card_image_list)):
        for _ in range(deck_list.iloc[i, 1]):
            deck_image_list.append(card_image_list[i])
    get_pdf_image(output_file_name, deck_image_list)
    
    
    