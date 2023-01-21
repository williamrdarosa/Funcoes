import telebot
from telebot import types
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton
from datetime import datetime
import os
from dotenv import load_dotenv, find_dotenv

def codigos(usuario):
    load_dotenv(find_dotenv())
    return os.environ.get(usuario)
    
def data_agora():
    return f'{datetime.now():%d/%m/%Y %H:%M}'

class Programa():

    def __init__(self, token):
        self.__token = token
        self.__bot = telebot.TeleBot(token, parse_mode='Markdown')

    def comando_mensagem(self, comando, mensagem):
        @self.__bot.message_handler(commands=[comando])
        def funcao_comando(message):
            self.__bot.reply_to(message, mensagem)

    def start(self, mensagem):
        @self.__bot.message_handler(commands=['start'])
        def funcao_start(message):
            self.__bot.reply_to(message, mensagem)
            chat = self.__bot.get_chat(message.chat.id)
            if message.chat.title is None:
                print(f'Requisição de Usuário')
                print(f'---------------------')
                print(f'ID do usuário: {message.chat.id}')
                print(f'Primeiro nome: {message.chat.first_name}')
                print(f'Primeiro segundo nome: {message.chat.last_name}')
                print(f'Nome de Usuário: {message.chat.username}')
                print(f'Bio de Usuário: {chat.bio}')
                print(f'---------------------')
            else:
                print(f'Requisição de Grupo')
                print(f'-------------------')
                print(f'ID do grupo: {message.chat.id}')
                print(f'Tipo do grupo: {message.chat.type}')
                print(f'Titulo do grupo: {message.chat.title}')
                print(f'Descrição do Grupo: {chat.description}')
                print(f'-------------------')

    def falar(self, id, mensagem):
        self.__bot.send_message(id, mensagem)

    def enviar_arquivo(self, id, arquivo):
        doc = open(arquivo, 'rb')
        self.__bot.send_document(id, doc)

    def enviar_imagem(self, id, imagem):
        photo = open(imagem, 'rb')
        self.__bot.send_photo(id, photo)

    def eco(self):
        @self.__bot.message_handler(func=lambda message: True)
        def funcao_eco(message):
            self.__bot.reply_to(message, message.text)

    def eco_mensagem(self,saudacao, mensagem):
        @self.__bot.message_handler(func=lambda message: True)
        def funcao_eco_mensagem(message):
            self.__bot.reply_to(message,f'{saudacao} {message.chat.first_name}, {mensagem}')

    def iniciar(self):
        self.__bot.infinity_polling()

    def receber_imagem(self):
        @self.__bot.message_handler(content_types=['photo'])
        def photo(message):
            fileID = message.photo[-1].file_id
            file_info = self.__bot.get_file(fileID)
            downloaded_file = self.__bot.download_file(file_info.file_path)

            with open("image.jpg", 'wb') as new_file:
                new_file.write(downloaded_file)

    def receber_arquivo(self):
        @self.__bot.message_handler(content_types=['document', 'photo', 'audio', 'video', 'voice']) # list relevant content types
        def addfile(message):
            file_name = message.document.file_name
            file_info = self.__bot.get_file(message.document.file_id)
            downloaded_file = self.__bot.download_file(file_info.file_path)
            with open(file_name, 'wb') as new_file:
                new_file.write(downloaded_file)