# -*- coding: utf-8 -*-
"""
Created on Tue Dec 29 02:01:06 2020

@author: olang
"""
import fitz
import os
import sys
import random
import shutil
from PIL import Image
import PySimpleGUI as sg

###############################################################################
def get_number_code(filename): # função que busca uma sequência de 7 números
    lst = list(filename)
    numbers = []
    output = []
    stringcode  = str()
    k = 0    
    for word in lst:
        if word.isdigit():
            k += 1
            numbers.append(int(word))
            if k == 7:
                output = numbers
        else:
            k = 0
            numbers = []
    for i in range(len(output)):
        stringcode = stringcode + str(output[i])
    return stringcode
###############################################################################
def rnd_gen(w,h,div):
    dw = round(w*div)
    dh = round(h*div)
    rndw = random.randint(1,dw)
    rndh = random.randint(1,dh)    
    rndW = random.randrange(1,w,rndw)
    rndH = random.randrange(1,h,rndh)    
    return [rndW,rndH]   
###############################################################################
def filter_img(pix,ratio,div):
    m = pix.width
    n = pix.height
    count = 0
    for i in range(round(m*div)):
        for j in range(round(n*div)):
            [a,b] = rnd_gen(m,n,0.04)
            px = pix.pixel(a,b)
            if px[0] == 76 and px[1] == 76 and px[2] == 76:
                count = count + 1
    p = count/((div*m)*(div*n))
    if p > ratio:
        return True
    else:
        return False
###############################################################################
def find_text(page,keyword):
    for j in range(len(keyword)):
        r = page.searchFor(keyword[j])
        if len(r)>0:
            if keyword[j] == keyword[0]:
                return "capa"
                break
            elif keyword[j] == keyword[1]:
                return "local"
                break
            elif keyword[j] == keyword[2]:
                return "relatorio_fotografico"
                break
            elif keyword[j] == keyword[3]:
                return "relacao_material"
                break
    return None
###############################################################################
def compare_pix(pix1,pix2):
    if pix1 != None and pix2 != None:
        if min(pix1.height,pix2.height) < 530 or 580 < min(pix1.height,pix2.height) < 600:
            if pix1.width == pix2.width:
                return True
            else:
                return False
        else:
            return None
    else:
        return None
###############################################################################
def get_concat_v(img1, img2, text):
    im1 = Image.open(text + '/' + img1)
    im2 = Image.open(img2)
    dst = Image.new('RGB', (im1.width, im1.height + im2.height))
    dst.paste(im2, (0, 0))
    dst.paste(im1, (0, im2.height))
    return dst
###############################################################################
def LoadLayout():
    layout1 = [[sg.Text("Extração das imagens terminada!")],[sg.Button("OK")]]
    
    return layout1
###############################################################################
def count_docs(directory):
    counter = 0
    for file in os.listdir(directory):
        if file.endswith(".pdf"):
            counter += 1
    return counter

###############################################################################
def main():
    if not tuple(map(int, fitz.version[0].split("."))) >= (1, 13, 17):
        raise SystemExit("require PyMuPDF v1.13.17+")
        
    directory = sys.argv[1] if len(sys.argv) == 2 else None
    if not directory:
        directory = sg.PopupGetFolder("Selecione a pasta:", title = "Extrator de Imagens")
    if not directory:
        raise SystemExit()
       
    dim_lim = 90
    ratio = 0.01
    keyword = ["Planejamento Executivo", "MAPA DE LOCALIZAÇÃO", "RELATÓRIO FOTOGRÁFICO", "RELAÇÃO MATERIAL"]
    m = 0
    text = " "
    pix = None
    num_docs = count_docs(directory)
    close = False
    finished = False
    for filename in os.listdir(directory): # Para cada arquivo no diretório
        if close == True:
            if finished == True:
                finished = False
            else:
                break
        if filename.endswith(".pdf"): # Se o arquivo tem o formato pdf              
            m = m+1
            k = 0
            doc = fitz.open(directory + '/' + filename)
            num = get_number_code(filename)
            num_pages = len(doc)
            for i in range(num_pages):
                meter = sg.OneLineProgressMeter("Extraindo Imagens", i+1, num_pages, '_M_','Páginas Escaneadas do documento %s de %s documento(s)' % (m,num_docs),orientation='h')
                if not meter:
                    close = True;
                    break
                
                page = doc.loadPage(i)
                text1 = text
                text = find_text(page,keyword)
                if text != None:          
                    if text != text1 and text1 != None:
                        k = 0
                        if not os.path.exists(text):
                            os.mkdir(text)                     
                    for img in doc.getPageImageList(i):
                        xref = img[0]
                        pix1 = pix
                        pix = fitz.Pixmap(doc,xref)                 
                        if pix.height > dim_lim and pix.width > dim_lim:
                            k = k+1
                            result = compare_pix(pix,pix1)
                            img_name = ("%s_%s_%s_%s.png" % (text,num,m,k))
                            img_name1 = ("%s_%s_%s_%s.png" % (text,num,m,k-1))
                            if pix.n < 5:
                                pix.writePNG(img_name)
                            else:
                                pix = fitz.Pixmap(fitz.csRGB,pix)
                                pix.writePNG(img_name)                    
                            cond = filter_img(pix,ratio,0.1)
                            if cond == True:
                                os.remove(img_name)
                                k = k-1
                            else:
                                if result == True:
                                    k = k-1
                                    im = get_concat_v(img_name1,img_name,text)
                                    os.remove(img_name)
                                    os.remove(text + '/' + img_name1)
                                    im.save(img_name1)
                                    shutil.move(img_name1,text + '/' + img_name1)
                                else:
                                    shutil.move(img_name,text + '/' + img_name)
                    else:
                        continue
            if i+1 == num_pages:
                finished = True
            else:
                finished = False
    window = sg.Window("Finalizado",LoadLayout())
    while True: 
        event, values = window.read()
        if event == "OK" or sg.WIN_CLOSED():
            break 
    window.close()

if __name__ == '__main__':
    main(sys.argv)
                        
                        