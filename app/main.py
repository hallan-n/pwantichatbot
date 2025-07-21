import re
import time
import pygetwindow as gw
import pyautogui
import pytesseract
from PIL import Image, ImageOps, ImageFilter
import os
import win32gui

# Configurações do Tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
os.environ['TESSDATA_PREFIX'] = r"C:\Program Files\Tesseract-OCR\tessdata"

def preprocess_image(img):
    rgb = img.convert('RGB')
    mask = Image.new('L', rgb.size, 0)
    pixels = rgb.load()
    mask_pixels = mask.load()
    width, height = rgb.size

    for y in range(height):
        for x in range(width):
            r, g, b = pixels[x, y]
            # filtra só azul forte para isolar o texto do chat
            if (
                b > 180 and       # azul alto
                r < 60 and        # vermelho baixo
                70 < g < 120      # verde moderado (mais que vermelho, menos que azul)
            ):
                mask_pixels[x, y] = 255
            else:
                mask_pixels[x, y] = 0


    processed = mask
    processed = ImageOps.autocontrast(processed)
    processed = processed.filter(ImageFilter.UnsharpMask(radius=2, percent=150, threshold=3))
    processed = processed.convert('L')
    return processed

def focar_janela(janela):
    hwnd = win32gui.FindWindow(None, janela.title)
    if hwnd:
        # traz a janela para frente sem redimensionar
        import win32con, win32api
        win32gui.SetForegroundWindow(hwnd)
        print(f"Focou a janela: {janela.title}")
    else:
        print(f"Janela não encontrada: {janela.title}")

def limpar_texto():
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.2)
    pyautogui.press('backspace')
    time.sleep(0.2)
    pyautogui.press('backspace')
    time.sleep(0.2)



def digitar_texto(texto):
    pyautogui.write(str(texto), interval=0.05)

def main():
    janelas = gw.getWindowsWithTitle('The Classic PW')
    if not janelas:
        print("Janela do PW não encontrada.")
        return

    janela = janelas[0]

    x, y, w, h = janela.left, janela.top, janela.width, janela.height

    region_x = x
    region_y = y + h // 2
    region_w = w // 2
    region_h = h // 2

    print(f"Capturando região: x={region_x}, y={region_y}, w={region_w}, h={region_h}")
    screenshot = pyautogui.screenshot(region=(region_x, region_y, region_w, region_h))
    screenshot.save("captura_original.png")

    img_processada = preprocess_image(screenshot)
    img_processada.save("captura_filtrada.png")

    try:
        texto = pytesseract.image_to_string(img_processada, lang='eng', config='--psm 6')
    except pytesseract.TesseractError as e:
        print(f"Erro no OCR: {e}")
        return

    print("\nTexto OCR completo:")
    print(texto)

    linhas = texto.splitlines()
    linha_quanto = ""
    for linha in linhas:
        if 'quanto' in linha.lower():
            linha_quanto = linha
            break

    if not linha_quanto:
        print("Linha com 'Quanto' não encontrada no OCR.")
        return

    linha_limpa = re.sub(r'.+Quanto\s{1,}.|.\s{1,}Res.+$', '', linha_quanto)

    numeros = re.findall(r'\d+', linha_limpa)
    numeros_int = list(map(int, numeros))
    soma = sum(numeros_int)

    print("\n=== Resultado ===")
    print(f"Linha original: {linha_quanto}")
    print(f"Linha após regex: {linha_limpa}")
    print(f"Números extraídos: {numeros_int}")
    print(f"Soma: {soma}")

    # Focar o PW e digitar resposta
    focar_janela(janela)
    time.sleep(0.5)

    # Abrir chat
    pyautogui.press('enter')
    time.sleep(0.3)

    # Limpar texto do chat
    limpar_texto()

    # Digitar soma
    digitar_texto(soma)
    time.sleep(0.3)

    # Enviar mensagem
    pyautogui.press('enter')

    print("Resposta digitada no chat!")

if __name__ == "__main__":
    main()
