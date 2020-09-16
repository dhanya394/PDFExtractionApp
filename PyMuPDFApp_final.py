import fitz
from statistics import mean
from collections import Counter
from itertools import groupby
import pandas as pd

#provide full path to PDF file
doc = fitz.open("E:/input_file.pdf")
pages = doc.pageCount #total page count

final_list = []

#Loop through all pages of the document
for i in range(0,pages):
    page = doc[i]
    blocks = page.getText("dict", flags=11)["blocks"] #Extract texts from pages as blocks
    for b in blocks:
        for l in b["lines"]: #looping through every textline of the page
            text_line = ''
            font_list = []
            font_size = []
            location = l["bbox"]
            list1 = []
            for s in l["spans"]:
                text_line = text_line + s["text"]
                font_list.append(s["font"])
                font_size.append(s["size"])
            freqs = groupby(Counter(font_list).most_common(), lambda x: x[1])
            mode = [val for val, count in next(freqs)[1]]
            list1.append(i)
            list1.append(text_line)
            list1.append(mode[0])
            list1.append(mean(font_size))
            list1.append(location)
            final_list.append(list1)


data = final_list #List containing final extracted data
#creating pandas Dataframe
df = pd.DataFrame(data,columns=['Page Number','Text Line', 'Font', 'Font Size', 'X,Y Coordinates of Text Line'],dtype=float)
print(df)
#converting Dataframe to Excel sheet
df.to_excel("E:/stuff/Barclays/PDF Extraction/output2.xlsx")
