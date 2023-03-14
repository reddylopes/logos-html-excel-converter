# Lib para ler HTML
from bs4 import BeautifulSoup

# Lib para time.sleep(1) no debug
import time

# Lib para dialog
import tkinter as tk
from tkinter import filedialog

# Lib para encerrar programa
import sys

# Abrir arquivo HTML
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
#file_path = "C:\\Users\\jonat\\Desktop\\jo7.html"

f = open(file_path, encoding="utf8")

# Parse do HTML com BeautifulSoup
html_doc = f.read()
soup = BeautifulSoup(html_doc, 'html.parser')
body = soup.find("body")

# Teste: se o HTML não estiver no formato padrão, encerrra
if len(body.find_all("h2")) > 1:
    print("Há um problema no seu HTML.\nÉ possível que você não esteja usando o guia exegético personalizado somente com a seção 'Palavra por palavra'.\nNo Logos, crie um novo guia, exclua todas as seções, deixando somente a seção 'Palavra por palavra' mostrando todas as partes do discurso.\nExporte novamente o HTML a partir desse novo guia personalizado e tente novamente.")
    input("Pressione Enter para encerrar...")
    sys.exit()

# Número de versículos
verse_count = len(body.find_all("blockquote", recursive=False))
print("Número de versos: ", verse_count, "\n")

# Encontra todos os versículos
verse_blockquotes = []
for bq in body.find_all('blockquote', recursive=False):
    verse_blockquotes.append(bq)
verse_divs = []
for div in body.find_all('div', recursive=False):
    style = div.get('style')
    if style is not None:
        if not "footer" in style:
            verse_divs.append(div)

# Lista os textos e referências
verse_refs = []
verse_texts = []
for verse in verse_divs:
    verse_ref = verse.table.tr.td.p.a.text
    print(verse_ref)
    verse_refs.append(verse_ref)
    paragraph = verse.table.find("tr").findNext("tr").td.p
    verse_text = paragraph.text
    verse_text = verse_text.split("|")[0].rstrip(verse_text[-4])
    verse_text = verse_text[1:]
    print(verse_text, "\n")
    verse_texts.append(verse_text)

# Palavras, morfologias, lemmas e significados
full_table=[]
# Iteração entre versos
i=0
for bq in verse_blockquotes:
    # Iteração entre palavras do verso
    for p in bq.find_all("p", recursive=False):
        # Palavra
        word = p.span.text
        # O loop aqui é necessário por causa do hebraico
        # No hebraico, uma mesma palavra pode conter prefixos e sufixos que também possuem lemmas, traduções e morfologias
        last = False
        analysis = ""
        word_feats=[]
        while last == False:
            # Próximos 3 blockquotes
            if word_feats == []:
                word_feats = p.find_next_siblings("blockquote", limit=3)
            else:
                word_feats = word_feats[2].find_next_siblings("blockquote", limit=3)
            # Lemma
            lemma = word_feats[0].find("a").span.text
            # Traduções
            translation = word_feats[0].find("span", attrs={'style':'font-weight:bold;'}).text
            # Morfologias
            morph = ""
            for span in word_feats[1].p.find_all('span'):
                morph = morph + span.text + " "
            morph = morph.capitalize().rstrip(morph[-1])
            analysis = analysis + morph + " (" + lemma + "): " + translation + "."
            if word_feats[2].find_next_sibling() is None:
                last = True
            elif word_feats[2].find_next_sibling().name != "blockquote":
                last = True
            else:
                analysis = analysis + "\n"
        # Preenche a tabela
        full_table.append([verse_refs[i], word, analysis])
    i += 1

print("Foram encontradas ", len(full_table), " palavras no arquivo.\n")

# Registros na planilha
import xlsxwriter

file_path = file_path.replace(".html", ".xlsx")
workbook = xlsxwriter.Workbook(file_path)
worksheet = workbook.add_worksheet()

# Registro dos textos dos versículos
for x in range(len(verse_refs)):
    worksheet.write(x, 0, verse_refs[x])
    worksheet.write(x, 1, verse_texts[x])

# Registro das palavras
worksheet.write(len(verse_ref), 0, "Ref")
worksheet.write(len(verse_ref), 1, "Manuscrito")
worksheet.write(len(verse_ref), 2, "Análise morfológica")
worksheet.write(len(verse_ref), 3, "Tradução")

for x in range(len(full_table)):
    worksheet.write(x+len(verse_ref)+1, 0, full_table[x][0])
    worksheet.write(x+len(verse_ref)+1, 1, full_table[x][1])
    worksheet.write(x+len(verse_ref)+1, 2, full_table[x][2])
     
workbook.close()

print("Planilha gerada com sucesso.\nCaminho: ", file_path, "\n\n")
intput("Pressione Enter para encerrar...")
