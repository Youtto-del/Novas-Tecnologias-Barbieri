import pytesseract.pytesseract
from PIL import Image
import pytesseract

# Imagem a ser quebrada, neste ponto você poderia usar urlib, httplib ou curl para carregar esta imagem.
img = Image.open('download.png')

# convertemos para o padrão RGB
img = img.convert("RGBA")

# damos bind da imagem para a variável pixdata
pixdata = img.load()

# Limpando a sujeira do background, se a cor for mais clara que a medida, então transformamos em branco,
# enquianto transformamos os pixels escuros em preto.
for y in range(img.size[1]):
    for x in range(img.size[0]):
        if sum(pixdata[x, y]) < 630:
            pixdata[x, y] = (0, 0, 0, 0)
        else:
            pixdata[x, y] = (255, 255, 255, 255)

# Salvamos a nova imagem com fundo branco
img.save("contraste.gif", "GIF")

#   Aumentamos as dimensões da imagem (requerido pelo OCR)
im_orig = Image.open('contraste.gif')
big = im_orig.resize((116, 56), Image.NEAREST)

# Salvamos a imagem com tamanho maior
ext = ".gif"
big.save("input-NEAREST" + ext)

# Yeah! Fazemos OCR da imagem usando o pytesseract

image = Image.open('input-NEAREST.gif')

# simplesmente imprimimos a imagem em formato de string OCRizado
print(pytesseract.image_to_string(image))
