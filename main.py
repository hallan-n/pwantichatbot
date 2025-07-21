import re
import time
import pygetwindow as gw
import pyautogui
import pytesseract
from PIL import Image, ImageOps, ImageFilter
import os
from pathlib import Path
import win32gui
import win32con
import win32api

BASE_DIR = Path(__file__).resolve().parent

# Config Tesseract interno
pytesseract.pytesseract.tesseract_cmd = str(BASE_DIR / "Tesseract-OCR" / "tesseract.exe")
os.environ['TESSDATA_PREFIX'] = str(BASE_DIR / "Tesseract-OCR" / "tessdata")


def preprocess_image(img):
    rgb = img.convert('RGB')
    mask = Image.new('L', rgb.size, 0)
    pixels = rgb.load()
    mask_pixels = mask.load()
    width, height = rgb.size

    for y in range(height):
        for x in range(width):
            r, g, b = pixels[x, y]
            # filtro azul royal restrito para evitar roxo
            if b > 180 and r < 60 and 70 < g < 120:
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
        # Verifica o estado da janela
        estado = win32gui.IsIconic(hwnd)  # True se minimizada
        if estado:
            # Só restaura se estiver minimizada
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.3)  # pequena pausa pra garantir restauração
        try:
            win32gui.SetForegroundWindow(hwnd)
            print(f"Focou a janela: {janela.title}")
        except Exception as e:
            print(f"Erro ao focar janela: {e}")
    else:
        print(f"Janela não encontrada: {janela.title}")

def limpar_texto():
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.2)
    for _ in range(5):
        pyautogui.press('backspace')
        time.sleep(0.1)


def digitar_texto(texto):
    pyautogui.write(str(texto), interval=0.05)


def enviar_enter(hwnd):
    # Envia WM_KEYDOWN e WM_KEYUP para tecla ENTER diretamente para a janela
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.05)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
    print("ENTER enviado via PostMessage")


def capturar_responder():
    janelas = gw.getWindowsWithTitle('The Classic PW')
    if not janelas:
        print("Janela do PW não encontrada.")
        return False

    janela = janelas[0]
    hwnd = win32gui.FindWindow(None, janela.title)
    if not hwnd:
        print("HWND da janela não encontrado.")
        return False

    x, y, w, h = janela.left, janela.top, janela.width, janela.height

    region_x = x
    region_y = y + h // 2
    region_w = w // 2
    region_h = h // 2

    screenshot = pyautogui.screenshot(region=(region_x, region_y, region_w, region_h))
    img_processada = preprocess_image(screenshot)

    try:
        texto = pytesseract.image_to_string(img_processada, lang='eng', config='--psm 6')
    except pytesseract.TesseractError as e:
        print(f"Erro no OCR: {e}")
        return False

    linhas = texto.splitlines()
    linha_quanto = ""
    for linha in linhas:
        if 'quanto' in linha.lower():
            linha_quanto = linha
            break

    if not linha_quanto:
        print("Linha com 'Quanto' não encontrada no OCR.")
        return False

    linha_limpa = re.sub(r'.+Quanto\s{0,}.|.\s{0,}Res.+$', '', linha_quanto)
    numeros = re.findall(r'\d+', linha_limpa)
    if not numeros:
        print("Nenhum número encontrado na expressão.")
        return False
    numeros_int = list(map(int, numeros))
    soma = sum(numeros_int)

    print(f"\nPergunta detectada: {linha_quanto}")
    print(f"Expressão limpa: {linha_limpa}")
    print(f"Números extraídos: {numeros_int}")
    print(f"Soma calculada: {soma}")

    # Focar e responder
    focar_janela(janela)
    time.sleep(1)  # Delay maior para garantir foco

    enviar_enter(hwnd)
    time.sleep(0.5)

    limpar_texto()
    digitar_texto(soma)
    time.sleep(0.3)

    enviar_enter(hwnd)
    time.sleep(0.3)

    # Limpar chat com Ctrl + D
    pyautogui.hotkey('ctrl', 'd')
    print("Resposta enviada e chat limpa com Ctrl+D!")

    return True


def main():
    print("Iniciando AntiBot automático. Pressione Ctrl+C para parar.")
    while True:
        try:
            respondeu = capturar_responder()
            if not respondeu:
                print("Nenhuma pergunta detectada. Tentando novamente em 5 segundos...")
            time.sleep(5)
        except KeyboardInterrupt:
            print("Encerrando script.")
            break
        except Exception as e:
            print(f"Erro inesperado: {e}")
            time.sleep(5)


if __name__ == "__main__":
    main()
