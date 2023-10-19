import telebot
from telebot import types
import openpyxl
from openpyxl.reader.excel import load_workbook
import time
import string
from background import keep_alive


#bot = telebot.TeleBot("5754400160:AAFvJI6-SIXzUnT5nXOQhWils_QwKAIyTj4")


@bot.message_handler(commands=["start"])
def start(message):
    #print("start")
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn1 = types.InlineKeyboardButton("Узнать расписание", callback_data="fac")
    markup.add(btn1)
    bot.send_message(message.chat.id, text="Добро пожаловать в РГРТУ!", reply_markup=markup)
