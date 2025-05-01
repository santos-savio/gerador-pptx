#v 2.27
import tkinter as tk
from tkinter import messagebox, colorchooser, filedialog, ttk
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

# Inicializa a variável path_img como None para evitar erros de referência antes da seleção
path_img = None

# Funções para a interface
def colar_conteudo():
    try:
        # Cola o conteúdo da área de transferência no campo de texto
        texto = janela.clipboard_get()
        campo_texto.delete(1.0, tk.END)  # Limpa o campo de texto
        campo_texto.insert(tk.END, texto)  # Insere o conteúdo colado
    except tk.TclError:
        messagebox.showerror("Erro", "Nada para colar!")

def selecionar_imagem():
    # Abre o seletor de arquivos para escolher uma imagem
    global path_img
    path_img = filedialog.askopenfilename(
    title="Selecionar Imagem",
    filetypes=(("Arquivos de Imagem", "*.png;*.jpg;*.jpeg"), ("Todos os arquivos", "*.*")))
    if path_img:
        label_caminho_imagem.config(text=f"Imagem: {path_img}")
    else:
        label_caminho_imagem.config(text="Nenhuma imagem selecionada")

def exibir_ajuda():
    # Cria uma nova janela com instruções de uso
    janela_ajuda = tk.Toplevel(janela)
    janela_ajuda.title("Ajuda")
    janela_ajuda.geometry("400x300")

    texto_ajuda = """
    

    Como usar o Gerador de Apresentações PPTX:

    1. Cole o conteúdo dos slides no campo de texto.
    2. Ajuste as configurações de layout e estilo.
    3. Selecione uma imagem de fundo (opcional).
    4. Clique em "Gerar PPTX" para criar a apresentação.

    Dicas:
    - Use o botão "Colar" para colar texto da área de transferência.
    - Escolha uma cor para o texto usando o seletor de cores.
    - Insira o nome da fonte e o tamanho desejados.
    """
    label_ajuda = tk.Label(janela_ajuda, text=texto_ajuda, justify=tk.LEFT)
    label_ajuda.pack(padx=10, pady=10)

def calcular_largura(altura):
    """Calcula a largura com base na altura para manter a proporção 16:9."""
    return (16 / 9) * altura

def hex_to_rgb(hex_color):
    """Converte uma cor hexadecimal para RGB."""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def criar_apresentacao():
    """Cria uma apresentação PowerPoint a partir do conteúdo do campo de texto."""
    global path_img

    # Coleta as configurações da interface
    altura_slide = float(altura_slide_entry.get())
    pos_x = float(pos_x_entry.get())
    pos_y = float(pos_y_entry.get())
    # largura_textbox = float(largura_textbox_entry.get())
    # altura_textbox = float(altura_textbox_entry.get())
    font_size = int(font_size_entry.get())
    font_name = font_name_entry.get()
    n_maximo = int(n_maximo_entry.get())
    r, g, b = hex_to_rgb(cor_selecionada.get())  # Converte a cor hexadecimal para RGB

    # Coleta o conteúdo do campo de texto
    linhas = campo_texto.get(1.0, tk.END).splitlines()
    if not linhas:
        messagebox.showwarning("Aviso", "O campo de texto está vazio!")
        return

    # Define o nome do arquivo com base na primeira linha do texto
    nome_arquivo = linhas[0].strip() + ".pptx"
    arquivo_ppt = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("Arquivos PowerPoint", "*.pptx")], initialfile=nome_arquivo)
    if not arquivo_ppt:
        return  # Usuário cancelou a operação

    # Cria a apresentação
    prs = Presentation()
    largura_slide = calcular_largura(altura_slide)
    prs.slide_width = Inches(largura_slide)
    prs.slide_height = Inches(altura_slide)

    # Verifica linhas grandes
    linhas_grandes = [i + 1 for i, linha in enumerate(linhas) if len(linha) > n_maximo]

    titulo = True
    for indice, linha in enumerate(linhas):
        linha = linha.rstrip()  # Remove espaços adicionais
        if not linha:
            continue  # Ignora linhas vazias

        slide = prs.slides.add_slide(prs.slide_layouts[5])  # Cria um slide branco e sem título

        # Adiciona a imagem de fundo, se selecionada
        if path_img:
            slide.shapes.add_picture(path_img, 0, 0, prs.slide_width, prs.slide_height)
        else:
            pass  # Se não houver imagem, não faz nada

        # Adiciona o textbox
        # textbox = slide.shapes.add_textbox(Inches(pos_x), Inches(pos_y), Inches(largura_textbox), Inches(altura_textbox))
        textbox = slide.shapes.add_textbox(Inches(pos_x), Inches(pos_y), 8, 2)
        text_frame = textbox.text_frame
        text_frame.text = linha
        text_frame.auto_size = True
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        p = text_frame.paragraphs[0]
        p.font.size = Pt(font_size)
        p.font.name = font_name
        p.font.color.rgb = RGBColor(r, g, b)
        p.alignment = PP_ALIGN.CENTER

        if titulo:
            p.font.size = Pt(font_size * 1.6)  # Aumenta o tamanho da fonte do título
            titulo = False

        # Adiciona um comentário no primeiro slide com as linhas grandes
        if indice == 0 and linhas_grandes:
            comentario = f"Linhas grandes (>{n_maximo} caracteres): {', '.join(map(str, linhas_grandes))}"
            slide.notes_slide.notes_text_frame.text = comentario

    # Salva a apresentação
    prs.save(arquivo_ppt)
    messagebox.showinfo("Sucesso", f"Apresentação salva como {arquivo_ppt}")

# Configuração da janela principal
janela = tk.Tk()
janela.title("Gerador de Apresentações PPTX")
janela.geometry("600x500")

# Frame para o campo de texto
frame_texto = tk.Frame(janela)
frame_texto.grid(row=0, column=0, padx=10, pady=10)

# Botão "Colar" acima do campo de texto
botao_colar = tk.Button(frame_texto, text="Colar", command=colar_conteudo)
botao_colar.grid(row=0, column=0, pady=5)

# Campo de texto para o conteúdo dos slides
campo_texto = tk.Text(frame_texto, width=40, height=20)
campo_texto.grid(row=1, column=0, pady=5)

# Notebook (TabControl) para organizar as configurações
notebook = ttk.Notebook(janela)
notebook.grid(row=0, column=1, padx=10, pady=10, sticky="n")

# Aba "Formato"
aba_formato = ttk.Frame(notebook)
notebook.add(aba_formato, text="Formato")

# Controles de configuração na aba "Formato"
tk.Label(aba_formato, text="Altura do Slide (polegadas):").grid(row=0, column=0, sticky="w")
altura_slide_entry = tk.Entry(aba_formato)
altura_slide_entry.grid(row=1, column=0, pady=2)
altura_slide_entry.insert(0, "7.5")

tk.Label(aba_formato, text="Posição X:").grid(row=2, column=0, sticky="w")
pos_x_entry = tk.Entry(aba_formato)
pos_x_entry.grid(row=3, column=0, pady=2)
# pos_x_entry.insert(0, "5")
pos_x_entry.insert(0, "6.65")

tk.Label(aba_formato, text="Posição Y:").grid(row=4, column=0, sticky="w")
pos_y_entry = tk.Entry(aba_formato)
pos_y_entry.grid(row=5, column=0, pady=2)
pos_y_entry.insert(0, "2")

# tk.Label(aba_formato, text="Largura da Caixa de Texto:").grid(row=6, column=0, sticky="w")
# largura_textbox_entry = tk.Entry(aba_formato)
# largura_textbox_entry.grid(row=7, column=0, pady=2)
# largura_textbox_entry.insert(0, "8")

# tk.Label(aba_formato, text="Altura da Caixa de Texto:").grid(row=8, column=0, sticky="w")
# altura_textbox_entry = tk.Entry(aba_formato)
# altura_textbox_entry.grid(row=9, column=0, pady=2)
# altura_textbox_entry.insert(0, "2")

tk.Label(aba_formato, text="Tamanho Máximo de Letras (nMaximo):").grid(row=10, column=0, sticky="w")
n_maximo_entry = tk.Entry(aba_formato)
n_maximo_entry.grid(row=11, column=0, pady=2)
n_maximo_entry.insert(0, "40")

# Aba "Texto"
aba_texto = ttk.Frame(notebook)
notebook.add(aba_texto, text="Texto")

# Controles de configuração na aba "Texto"
tk.Label(aba_texto, text="Tamanho da Fonte:").grid(row=0, column=0, sticky="w")
font_size_entry = tk.Entry(aba_texto)
font_size_entry.grid(row=1, column=0, pady=2)
font_size_entry.insert(0, "48")

tk.Label(aba_texto, text="Nome da Fonte:").grid(row=2, column=0, sticky="w")
font_name_entry = tk.Entry(aba_texto)
font_name_entry.grid(row=3, column=0, pady=2)
font_name_entry.insert(0, "BANDEX")

# Seletor de cor
def escolher_cor():
    cor = colorchooser.askcolor()[1]  # Retorna o código hexadecimal da cor
    if cor:
        cor_selecionada.set(cor)  # Armazena a cor selecionada
        botao_cor.config(bg=cor)

cor_selecionada = tk.StringVar(value="#FFFFFF")  # Cor padrão: branco
tk.Label(aba_texto, text="Cor do Texto:").grid(row=4, column=0, sticky="w")
botao_cor = tk.Button(aba_texto, text="Escolher Cor", command=escolher_cor, bg="white")
botao_cor.grid(row=5, column=0, pady=2)

# Define a aba "Texto" como padrão
notebook.select(aba_texto)

# Frame para os botões
frame_botoes = tk.Frame(janela)
frame_botoes.grid(row=1, column=0, columnspan=2, pady=10)

# Botão "Selecionar Imagem"
botao_imagem = tk.Button(frame_botoes, text="Selecionar Imagem", command=selecionar_imagem)
botao_imagem.grid(row=0, column=0, padx=5)

# Rótulo para exibir o caminho da imagem
label_caminho_imagem = tk.Label(frame_botoes, text="Nenhuma imagem selecionada", fg="gray", wraplength=300)  # Define a largura máxima em pixels para o texto
label_caminho_imagem.grid(row=1, column=0, padx=5, pady=2)

# Botão "Gerar PPTX"
botao_gerar = tk.Button(frame_botoes, text="Gerar PPTX", command=criar_apresentacao)
botao_gerar.grid(row=0, column=1, padx=5)

# Botão "Ajuda"
botao_ajuda = tk.Button(frame_botoes, text="Ajuda", command=exibir_ajuda)
botao_ajuda.grid(row=0, column=2, padx=5)

# Inicia a interface
janela.mainloop()
