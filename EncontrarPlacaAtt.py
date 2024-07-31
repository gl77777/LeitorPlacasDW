import cv2
import pytesseract
from PIL import Image, ImageTk
import customtkinter as ctk
from tkinter import Label
import pandas as pd
import json
import difflib
import time
from comtypes.client import CreateObject

class PyE3DataAccess(object):
    def __init__(self, server='192.168.10.10'):
        super(PyE3DataAccess, self).__init__()
        self._engine = CreateObject("{80327130-FFDB-4506-B160-B9F8DB32DFB2}")
        self._engine.Server = server

    def lerValorE3(self, pathname):
        return self._engine.ReadValue(pathname)
    
    def escreverValorE3(self, pathname, date, quality, value):
        return self._engine.WriteValue(pathname, date, quality, value)


# Ler o arquivo JSON
with open("dados.json", 'r', encoding='utf-8') as file:
    data = json.load(file)

# Transformar o JSON em um DataFrame
df = pd.DataFrame(data["items"])

# Configuração do Tesseract
pytesseract.pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

class WebcamApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.lista_placas = list(df['placa_caminhao'])

        self.title("Webcam App")
        self.geometry("1000x800")
        self.configure(fg_color="#232322")
        my_image = ctk.CTkImage(dark_image=Image.open("datawake.png"), size=(500, 100))
        self.image_label = ctk.CTkLabel(self, image=my_image, text="")  # display image with a CTkLabel
        self.image_label.pack(pady=10, padx=10)
        

        self.video_label = Label(self)
        self.video_label.pack(padx=10, pady=10)

        self.resultado_label = ctk.CTkLabel(self, text="", font=("Arial", 20))
        self.resultado_label.pack(pady=20, padx=10)

        self.btn = ctk.CTkButton(self, text="Concluir", command=self.concluir, fg_color="#FDB913", text_color="black")
        self.btn.pack(pady=10, padx=10)
        self.btn2 = ctk.CTkButton(self, text="Recusar", command=self.recusado, fg_color="#FDB913", text_color="black")
    

        self.cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
        if not self.cap.isOpened():
            print("Erro ao abrir a câmera.")
            return

        self.update_video()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def desenhaContornoMaiorArea(self, contornos, imagem):
        maior_area = 0
        melhor_contorno = None

        for c in contornos:
            area = cv2.contourArea(c)
            if area > maior_area:
                maior_area = area
                melhor_contorno = c
        
        if melhor_contorno is not None:
            perimetro = cv2.arcLength(melhor_contorno, True)
            approx = cv2.approxPolyDP(melhor_contorno, 0.03 * perimetro, True)
            if len(approx) == 4:
                (x, y, lar, alt) = cv2.boundingRect(melhor_contorno)
                cv2.rectangle(imagem, (x, y), (x + lar, y + alt), (0, 255, 0), 3)
                roi = imagem[y:y + alt, x:x + lar]
                cv2.imwrite("output/roi.png", roi)

    def preProcessamentoRoi(self):
        # Carrega a imagem da ROI
        img_roi = cv2.imread("output/roi.png")
        if img_roi is None:
            print("Erro ao carregar imagem da ROI.")
            return None

        # Redimensiona a imagem
        img = cv2.resize(img_roi, None, fx=4, fy=4, interpolation=cv2.INTER_CUBIC)

        # Converte para escala de cinza
        img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # Aplica limiarização
        _, img = cv2.threshold(img, 70, 255, cv2.THRESH_BINARY)

        # Aplica desfoque gaussiano
        img = cv2.GaussianBlur(img, (5, 5), 0)

        # Salva a imagem processada para OCR
        cv2.imwrite("output/roi-ocr.png", img)

        return self.reconhecimentoOCR()

    def reconhecimentoOCR(self):
        # Carrega a imagem processada para OCR
        img_roi_ocr = cv2.imread("output/roi-ocr.png")
        if img_roi_ocr is None:
            print("Erro ao carregar imagem para OCR.")
            return None

        # Configuração para o Tesseract
        config = r'-c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 --psm 6'
        
        # Realiza OCR na imagem e exibe o resultado
        saida = pytesseract.image_to_string(img_roi_ocr, lang='eng', config=config)
        
        # Limita o resultado para 7 caracteres
        resultado_limitado = saida[:7]
        
        print(resultado_limitado)
        return resultado_limitado

    def update_video(self):
        ret, frame = self.cap.read()
        if ret:
            height, width = frame.shape[:2]
            top_left_x = int(width * 0.25)
            bottom_right_x = int(width * 0.75)
            top_left_y = int(height * 0.7)

            area = frame[top_left_y:height, top_left_x:bottom_right_x]

            if area.size != 0:
                # Processamento da área para melhorar visualização
                img_result = cv2.cvtColor(area, cv2.COLOR_BGR2GRAY)
                _, img_result = cv2.threshold(img_result, 90, 255, cv2.THRESH_BINARY)
                img_result = cv2.GaussianBlur(img_result, (5, 5), 0)

                contornos, hier = cv2.findContours(img_result, cv2.RETR_TREE, cv2.CHAIN_APPROX_NONE)
                self.desenhaContornoMaiorArea(contornos, area)

            cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGBA)
            img = Image.fromarray(cv2image)
            imgtk = ImageTk.PhotoImage(image=img)
            self.video_label.imgtk = imgtk
            self.video_label.configure(image=imgtk)
        self.after(10, self.update_video)

    def concluir(self):
        self.resultado = self.preProcessamentoRoi()
        
        index = 0
        dfResultado = df[['nomefornecedor','descricaoproduto']]
        produtos = []
        if self.resultado  in self.lista_placas:
            for i in self.lista_placas:
                if self.resultado  == i:
                    produtos.append(dfResultado.loc[index, 'descricaoproduto'])
            self.resultado_label.configure(text=f"Resultado OCR: {self.resultado } \nFornecedor: {dfResultado.loc[index, 'nomefornecedor']}")
            self.btn.configure(text='Confirmar', command=self.liberado)
            self.btn2.pack(pady=10, padx=10)
            
                    
                    
                
                
        if self.resultado not in self.lista_placas:
            self.resultado_label.configure(text=f"Resultado OCR: {self.resultado} \nDADOS NÃO ENCONTRADOS")

        if self.resultado  == '':
            self.resultado_label.configure(text=f"Resultado: DADOS NÃO ENCONTRADOS")
    def liberado(self):
        self.resultado_label.configure(text=f"LIBERADO")
        valorTag = self.resultado 

        # Identificação para a conexão com o elipse
        pyE3DataAccess = PyE3DataAccess(server="192.168.10.10")
        caminhocaminhao = 'DdoInspRece.Geral.PlacaCaminhao.value'
        caminhook = 'DdoInspRece.Geral.OK.value'
        date = time.strftime("%d-%m-%Y %H:%M:%S", time.gmtime())
    
        pyE3DataAccess.escreverValorE3(pathname= caminhocaminhao, date=date, quality=192, value = valorTag)
        pyE3DataAccess.escreverValorE3(pathname= caminhook, date=date, quality=192, value = True)
        self.btn2.configure(state='disable')
        self.btn2.pack_forget()
        self.btn.configure(text='Voltar para tela inicial', command=self.clear)

    def recusado(self):
        self.resultado_label.configure(text=f"RECUSADO")
        self.btn2.configure(state='disable')
        self.btn2.pack_forget()
        self.btn.configure(text='Voltar para tela inicial', command=self.clear)



    def on_closing(self):
        self.cap.release()
        self.destroy()

    def clear(self):
        self.resultado_label.configure(text=f"")
        self.btn.configure(text="Concluir", command=self.concluir)
        self.btn2.configure(state='disable')



if __name__ == "__main__":
    app = WebcamApp()
    app.mainloop()


