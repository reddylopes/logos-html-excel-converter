from bs4 import BeautifulSoup
import time

import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename()
#file_path = "C:\\Users\\jonat\\Desktop\\jo7.html"

f = open(file_path, encoding="utf8")
html_doc = f.read()

soup = BeautifulSoup(html_doc, 'html.parser')

####################################
# Pegar todas as palavras do verso #
####################################
verse_words=[]

allps = soup.find_all('p')

for p in allps:
    style = p.get('style')
    if style is not None:
        if "background-color:rgb(215, 215, 215); margin-top:2pt; margin-bottom:2pt;" in style:
            verse_words.append(p.span.text)

##############################
# Pegar todas as morfologias #
##############################
morphs=[]

for p in allps:
    morph = ""
    if p.a is not None:
        href = p.a.get("href")
        if href is not None:
            if "https://ref.ly/logosref/morph-field" in href:
                for span in p.find_all('span'):
                    morph = morph + span.text + " "
                morphs.append(morph.capitalize().rstrip(morph[-1]))

#########################
# Pegar todos os lemmas #
#########################
lemmas=[]

allas = soup.find_all('a')

for a in allas:
    href = a.get("href")
    if href is not None:
        if "https://ref.ly/logos4/Guide?lemma=" in href:
            lemmas.append(a.span.text)

###############################
# Pegar todos os significados #
###############################
meanings=[]

allspans = soup.find_all('span')

for span in allspans:
    style = span.get("style")
    par1 = span.parent
    par2 = par1.parent
    if style is not None:
        if "font-weight:bold;" in style:
            if "p" in par1.name:
                if "td" in par2.name:
                    meanings.append(span.text)

print("Palavras:", len(verse_words))
print("Análises:", len(morphs))
print("  Lemmas:", len(lemmas))
print("Sentidos:", len(meanings))

##import csv
##
##if(len(verse_words) == len(morphs) == len(lemmas) == len(meanings)):
##    with open(file_path.replace(".html", ".csv"), 'w', encoding="utf-32", newline='') as file:
##        writer = csv.writer(file, delimiter = "|")
##        writer.writerow(["Manuscrito", "Análise morfológica", "Tradução"])
##        for x in range(len(verse_words)):
##            writer.writerow([verse_words[x], morphs[x] + " (" + lemmas[x] + "): " + meanings[x], ""])

import xlsxwriter

workbook = xlsxwriter.Workbook(file_path.replace(".html", ".xlsx"))
worksheet = workbook.add_worksheet()

worksheet.write(0, 0, "Manuscrito")
worksheet.write(0, 1, "Análise morfológica")
worksheet.write(0, 2, "Tradução")

for x in range(len(verse_words)):
    worksheet.write(x+1, 0, verse_words[x])
    worksheet.write(x+1, 1, morphs[x] + " (" + lemmas[x] + "): " + meanings[x], "")
     
workbook.close()

print("Deu bom")
