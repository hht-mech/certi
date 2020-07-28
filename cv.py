from PIL import Image, ImageDraw, ImageFont

import pandas as panda

# import the file that contains the name of participants
data = panda.read_excel("name_list.xlsx")

# import 'Full Name' List from the imported .xlsx file
names = data['Full Name'].to_list()

# determining the length of the list
max_no = len(names)

# iterate for creating certificates in bulk
# for i in names:
for index, name in enumerate(names, start = 1):
    image = Image.open("template.png")
    #width, height = image.size
    #print(width, height)
    draw = ImageDraw.Draw(image)
    W, H = (3900, 3145)
    # text color in RGB values
    text_color = (43, 61, 136)
    font = ImageFont.truetype("fonts/baskervi.ttf", 180, encoding="unic")
    w, h = draw.textsize(name.title(), font=font)
    location = ((W-w)/2,(H-h)/2)
    draw.text(location, name.title(), fill=text_color, font=font)
    image.save("generated/" + name + ".jpg")
    print(f"({index}/{max_no}) Certificate Created for {name.title()}")
    #print(i, index) => i indicates Name, index indicates its index position (start from 0 by default)