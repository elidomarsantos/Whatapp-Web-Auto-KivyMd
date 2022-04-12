from dataclasses import replace
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import urllib
from kivymd.app import MDApp
from kivy.lang import Builder
from kivy.core.window import Window
import sqlite3
from webdriver_manager.chrome import ChromeDriverManager
import os, sys
from kivy.resources import resource_add_path, resource_find
from kivymd.uix.button import MDFlatButton
from kivymd.uix.snackbar import BaseSnackbar
from kivy.properties import StringProperty, NumericProperty


# App para automatizar mensagens no Whatsapp Web, com interface gráfica construída com KivyMd, e utilizando SQLite.
# Necessário baixar o "Selenium Chromedriver" (se utilizar o Chrome), na versão correspondente ao Chrome,
# e colocar na pasta de instalação do Python

# Necessário também criação de uma pasta (C:/WA), para arquivamento das planilhas baixadas, e também para
# colocar os arquivos anexos em imagem, pdf, planilhas ou documentos, sempre seguindo como padrão
# ('nome_do arquivo'. (jpg, pdf,xlsx,docx), de acordo com o anexo correspondente

# Tamanho da tela
Window.size = (980, 660)

# Conexão (também criando, se não existir) ao banco de dados
conn = sqlite3.connect("lista_contatos.db")

class CustomSnackbar(BaseSnackbar):
    text = StringProperty(None)
    icon = StringProperty(None)
    font_size = NumericProperty("15sp")
    

class Whatsauto(MDApp):
    # Construtor do app
    def build(self):
        self.theme_cls.material_style = "M3"
        self.theme_cls.primary_palette = "Blue"
        self.theme_cls.accent_palette = "Green"
        return Builder.load_file('wapp.kv')

    # Para exportar uma planilha excel contendo contatos, para o banco de dados. 
    # Para exportar todos os contatos, pode-se usar o Google Contatos. Arrumar a planilha para conter apenas duas colunas> 'nome' e 'telefone', e nome 'contatos.xlsx'
    # Se já houver dados no BD, primeiro exportá-los (função abaixo) e incluí-los na planilha a ser exportada
    def sql_para_excel(self):
        contatos = pd.read_excel('C:/WA/contatos.xlsx')   
        contatos.to_sql('contatos', con=conn, if_exists='replace', index=False) 
        conn.commit()

    #A seguir as várias funções, com nomes autoexplicativos
    def inserir_contatos(self):
        nome = self.root.ids.inserir_nome.text
        nome_maius = nome.upper()
        telefone = self.root.ids.inserir_telefone.text
        telefone_completo = '5571'+telefone
        conn.execute('''INSERT INTO contatos(nome, telefone)
                         VALUES(?,?)
        ''',(nome_maius,telefone_completo))
        
        conn.commit()
        

    def remover_contatos_nome(self):
        try:
            remover = self.root.ids.de_remove.text
            remover_maius = remover.upper()
            conn.execute("""
            DELETE FROM contatos
            WHERE nome = ?
            """, (remover_maius))
            print
            conn.commit()

        except:
            snackbar = CustomSnackbar(
            text = "Este Contato não Existe!!",
            icon="information",
            snackbar_x="10dp",
            snackbar_y="10dp",
    
        )
        snackbar.size_hint_x = (
            Window.width - (snackbar.snackbar_x * 2)
        ) / Window.width
        snackbar.open()
 
    def remover_contatos_tel(self):
        try:
            remover_tel = self.root.ids.de_remove_tel.text
            remover_tel_completo = '5571'+remover_tel
            conn.execute("""
            DELETE FROM contatos
            WHERE telefone = ?
            """, (remover_tel_completo))

            conn.commit()
        
        except:
            snackbar = CustomSnackbar(
            text = "Este Contato não Existe!!",
            icon="information",
            snackbar_x="10dp",
            snackbar_y="10dp",
    
        )
        snackbar.size_hint_x = (
            Window.width - (snackbar.snackbar_x * 2)
        ) / Window.width
        snackbar.open()
            

    def remover_contatos_numero(self):
        
        de_ = self.root.ids.de_.text
        ate_ = self.root.ids.ate_.text
        conn.execute("""
        DELETE FROM contatos
        WHERE ROWID BETWEEN ? AND ? 
        """, (de_, ate_)).fetchall()
        

        conn.commit()

    def remover_contatos_all(self):
       
        conn.execute("""
        DELETE from contatos
        """, )
        
        conn.commit()
          

    def envio_sem_anexo(self):
        
        navegador = webdriver.Chrome()
        navegador.get('https://web.whatsapp.com/')
        time.sleep(30)
        
        de = self.root.ids.inicio.text
        ate = self.root.ids.fim.text
        c = conn.cursor()
        send = c.execute("""
        SELECT * FROM contatos WHERE ROWID BETWEEN ? AND ? 
        """, (de, ate)).fetchall()
     
        
        mensagem = self.root.ids.mensagem.text
        print(mensagem)
        for numero in send:
            
            print(numero[1])
            texto = urllib.parse.quote(f'{mensagem}')
            link = f'https://web.whatsapp.com/send?phone={numero}&text={texto}'
            navegador.get(link)
            time.sleep(20)
            navegador.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2]').send_keys(Keys.ENTER)
            time.sleep(1) 
        
        snackbar = CustomSnackbar(
            text = "Mensagens enviadas com sucesso!!",
            icon="information",
            snackbar_x="10dp",
            snackbar_y="10dp",
    
        )
        snackbar.size_hint_x = (
            Window.width - (snackbar.snackbar_x * 2)
        ) / Window.width
        snackbar.open()
       

    def envio_com_imagem(self):
        navegador = webdriver.Chrome()
        navegador.get('https://web.whatsapp.com/')
        time.sleep(30)
        
        de = self.root.ids.inicio.text
        ate = self.root.ids.fim.text
        c = conn.cursor()
        send = c.execute("""
        SELECT * FROM contatos WHERE ROWID BETWEEN ? AND ? 
        """, (de, ate)).fetchall()
     
        
        mensagem = self.root.ids.mensagem.text
        print(mensagem)
        for numero in send:
            
            print(numero[1])
            texto = urllib.parse.quote(f'{mensagem}')
            link = f'https://web.whatsapp.com/send?phone={numero}&text={texto}'
            navegador.get(link)
            time.sleep(20)
            navegador.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2]').send_keys(Keys.ENTER)
            time.sleep(5) 
            
            navegador.find_element_by_css_selector("span[data-icon='clip']").click()
            attach = navegador.find_element_by_css_selector("input[type='file']")
            attach.send_keys('C:/WA/1.jpg')
            time.sleep(5)
            send_file = navegador.find_element_by_css_selector("span[data-icon='send']")
            send_file.click()
            time.sleep(1)  

        snackbar = CustomSnackbar(
            text = "Mensagens enviadas com sucesso!!",
            icon="information",
            snackbar_x="10dp",
            snackbar_y="10dp",
    
        )
        snackbar.size_hint_x = (
            Window.width - (snackbar.snackbar_x * 2)
        ) / Window.width
        snackbar.open()


    def envio_com_pdf(self):
        navegador = webdriver.Chrome()
        navegador.get('https://web.whatsapp.com/')
        time.sleep(30)
        
        de = self.root.ids.inicio.text
        ate = self.root.ids.fim.text
        c = conn.cursor()
        send = c.execute("""
        SELECT * FROM contatos WHERE ROWID BETWEEN ? AND ? 
        """, (de, ate)).fetchall()
     
        
        mensagem = self.root.ids.mensagem.text
        print(mensagem)
        for numero in send:
            
            print(numero[1])
            texto = urllib.parse.quote(f'{mensagem}')
            link = f'https://web.whatsapp.com/send?phone={numero}&text={texto}'
            navegador.get(link)
            time.sleep(20)
            navegador.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2]').send_keys(Keys.ENTER)
            time.sleep(5) 
            
            navegador.find_element_by_css_selector("span[data-icon='clip']").click()
            attach = navegador.find_element_by_css_selector("input[type='file']")
            attach.send_keys('C:/WA/1.pdf')
            time.sleep(5)
            send_file = navegador.find_element_by_css_selector("span[data-icon='send']")
            send_file.click()
            time.sleep(1)  

        snackbar = CustomSnackbar(
            text = "Mensagens enviadas com sucesso!!",
            icon="information",
            snackbar_x="10dp",
            snackbar_y="10dp",
    
        )
        snackbar.size_hint_x = (
            Window.width - (snackbar.snackbar_x * 2)
        ) / Window.width
        snackbar.open()

    def envio_com_excel(self):
        navegador = webdriver.Chrome()
        navegador.get('https://web.whatsapp.com/')
        time.sleep(30)
        
        de = self.root.ids.inicio.text
        ate = self.root.ids.fim.text
        c = conn.cursor()
        send = c.execute("""
        SELECT * FROM contatos WHERE ROWID BETWEEN ? AND ? 
        """, (de, ate)).fetchall()
     
        
        mensagem = self.root.ids.mensagem.text
        print(mensagem)
        for numero in send:
            
            print(numero[1])
            texto = urllib.parse.quote(f'{mensagem}')
            link = f'https://web.whatsapp.com/send?phone={numero}&text={texto}'
            navegador.get(link)
            time.sleep(20)
            navegador.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2]').send_keys(Keys.ENTER)
            time.sleep(1) 
            
            navegador.find_element_by_css_selector("span[data-icon='clip']").click()
            attach = navegador.find_element_by_css_selector("input[type='file']")
            attach.send_keys('C:/WA/1.xlsx')
            time.sleep(5)
            send_file = navegador.find_element_by_css_selector("span[data-icon='send']")
            send_file.click()
            time.sleep(1)  

        snackbar = CustomSnackbar(
            text = "Mensagens enviadas com sucesso!!",
            icon="information",
            snackbar_x="10dp",
            snackbar_y="10dp",
    
        )
        snackbar.size_hint_x = (
            Window.width - (snackbar.snackbar_x * 2)
        ) / Window.width
        snackbar.open() 

    def envio_com_word(self):
        navegador = webdriver.Chrome()
        navegador.get('https://web.whatsapp.com/')
        time.sleep(30)
        
        de = self.root.ids.inicio.text
        ate = self.root.ids.fim.text
        c = conn.cursor()
        send = c.execute("""
        SELECT * FROM contatos WHERE ROWID BETWEEN ? AND ? 
        """, (de, ate)).fetchall()
     
        
        mensagem = self.root.ids.mensagem.text
        print(mensagem)
        for numero in send:
            
            print(numero[1])
            texto = urllib.parse.quote(f'{mensagem}')
            link = f'https://web.whatsapp.com/send?phone={numero}&text={texto}'
            navegador.get(link)
            time.sleep(20)
            navegador.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2]').send_keys(Keys.ENTER)
            time.sleep(5) 
            
            navegador.find_element_by_css_selector("span[data-icon='clip']").click()
            attach = navegador.find_element_by_css_selector("input[type='file']")
            attach.send_keys('C:/WA/1.docx')
            time.sleep(5)
            send_file = navegador.find_element_by_css_selector("span[data-icon='send']")
            send_file.click()
            time.sleep(1)  

        snackbar = CustomSnackbar(
            text = "Mensagens enviadas com sucesso!!",
            icon="information",
            snackbar_x="10dp",
            snackbar_y="10dp",
    
        )
        snackbar.size_hint_x = (
            Window.width - (snackbar.snackbar_x * 2)
        ) / Window.width
        snackbar.open()           

    # Função para atualizar a quantidade da contatos, a cada inserção/exclusão
    def q_contatos(self):
        c = conn.cursor()
        count = c.execute('''
        SELECT count(ROWID) FROM contatos
        ''').fetchall()
        c2 = str(count)[2:-3]
        self.root.ids.quant_contatos.text = (f'{c2}')

    # Mesma funcionalidade do anterior, para visualização da quantidade na abertura do app    
    def on_start(self):
        c = conn.cursor()
        count = c.execute('''
        SELECT count(ROWID) FROM contatos
        ''').fetchall()
        c2 = str(count)[2:-3]
        self.root.ids.quant_contatos.text = (f'{c2}')

    # Exporta a lista de contatos para um .csv, salvando em pasta pré-definida
    def baixar_em_excel(self):
        c = conn.cursor()
        count = c.execute('''
        SELECT * FROM contatos
        ''')
        print(count)
        df = pd.DataFrame(count)
        df.to_csv (r'C:\WA\lista_planilha.csv', index=False)
    
    # Fecha a conexão com SQLite, ao encerrar o app
    def on_stop(self):
        conn.close()

# Para compilar em um executável, com pyinstaller
if __name__ == '__main__':
    if hasattr(sys, '_MEIPASS'):
        resource_add_path(os.path.join(sys._MEIPASS))
    Whatsauto().run()

#pyinstaller --onefile --noconsole -w wapp.spec wapp.py