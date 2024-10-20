#


import tkinter as tk
from tkinter import filedialog
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

# Defina a resolução da tela
largura_tela = 1920 
altura_tela = 1080

pos_x = 1     # Posição X
pos_y = 2     # Posição Y
largura = 8   # Largura textbox
altura = 2    # Altura textbox
font_size = 32  # Tamanho da fonte
font_name = "BANDEX" # Fonte
cor = [255, 255, 255]


def selecionar_arquivo():
    """Abre um seletor de arquivo para escolher arquivos .txt."""
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal do tkinter
    arquivos_txt = filedialog.askopenfilenames(
        title="Selecione os arquivos .txt",
        filetypes=[("Arquivo de Texto", "*.txt")]
    )
    return arquivos_txt  # Retorna uma tupla com os caminhos dos arquivos

def calcular_polegadas(EMU):
    return EMU / 914400

def criar_apresentacao(txt_file, pos_x, pos_y, largura, altura, font_size, font_name, r, g, b):
    """Cria uma apresentação PowerPoint a partir de um arquivo .txt."""
    # Define o nome do arquivo PowerPoint de saída com base no nome do arquivo de texto
    arquivo_ppt = f'{txt_file.split("/")[-1].replace(".txt", "")}.pptx'
    
    prs = Presentation()

    print("A largura da apresentação é: ", prs.slide_width)

    # Abre o arquivo .txt e lê as linhas
    with open(txt_file, 'r', encoding='utf-8') as file:
        linhas = file.readlines()


    # Para cada linha do arquivo .txt, cria um slide e adiciona o texto
    for linha in linhas:
        linha = linha.strip()  # Remove espaços e quebras de linha adicionais
        print(linha)
        if linha:  # Só adiciona se a linha não for vazia
            slide = prs.slides.add_slide(prs.slide_layouts[5])  # Cria um slide branco e sem título
            #adiciona o textbox:
            textbox = slide.shapes.add_textbox(Inches(pos_x), Inches(pos_y), Inches(largura), Inches(altura))
            text_frame = textbox.text_frame # Acessa o textbox
            text_frame.text = linha # Insere o conteúdo da string linha no textbox
            text_frame.auto_size = True    # Ajusta automaticamente o tamanho da caixa
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE  # Centraliza verticalmente
            p = text_frame.paragraphs[0] # Acessa o primeiro (e único) parágrafo
            p.font.size = Pt(font_size) # Define o tamanho da fonte
            p.font.name = font_name # Define a fonte
            p.font.color.rgb = RGBColor(r, g, b)
            p.alignment = PP_ALIGN.CENTER

    path_img = add_imagem()
    print("path_img: ", path_img)

    # Itera sobre cada slide
    for slide in prs.slides:
        # Verifica se o slide tem um título
        if slide.shapes.title:
            # Verifica se o título está vazio
            if not slide.shapes.title.text.strip():
                # Remove a caixa de título se estiver vazia
                slide.shapes._spTree.remove(slide.shapes.title._element)
        slide.shapes.add_picture(path_img, 0, 0, largura_tela / 9, altura_tela / 9)

    # Salva a apresentação PowerPoint
    prs.save(arquivo_ppt)
    print(f'Apresentação {arquivo_ppt} criada com sucesso!')

def add_imagem():
    janela_add_img = tk.Toplevel()
    janela_add_img.withdraw()  # Esconde a janela principal do tkinter
    arquivo_img = filedialog.askopenfilenames(
        title="Selecione o arquivo .jpg ou png para fundo do slide",
        filetypes=[
            ("Arquivo de Imagem", "*.jpg"),
            ("Arquivo de Imagem", "*.png")]
    )
    
    if arquivo_img:
        return arquivo_img[0]  # Retorna apenas o primeiro arquivo selecionado
    else:
        return None  # Caso nenhum arquivo seja selecionado

largura_polegadas = calcular_polegadas(largura_tela)
altura_polegadas = calcular_polegadas(altura_tela)
r = cor[0]
g = cor[1]
b = cor[2]

# Seleciona os arquivos .txt
arquivos_txt = selecionar_arquivo()

# Cria a apresentação PowerPoint
if arquivos_txt:
    
    for arquivo_txt in arquivos_txt:  # Itera sobre cada arquivo selecionado
        criar_apresentacao(arquivo_txt, pos_x, pos_y, largura, altura, font_size, font_name, r, g, b)
