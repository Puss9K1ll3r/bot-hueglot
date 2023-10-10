from tokenize import group
from background import keep_alive
import pip

pip.main(['install', 'pytelegrambotapi'])
import telebot
from telebot import types
import os
import time
from zipfile import ZipFile
import openpyxl
import db
import sqlite3
import ScheduleRedactor as ShedRed

Dick<String, Dick<String, String>>
# бот вступает в работу

dictForCallbackData = {}
dictForDays = {
  "monday": "Понедельник",
  "tuesday": "Вторник",
  "wednesday": "Среда",
  "thursday": "Четверг",
  "friday": "Пятница",
  "saturday": "Суббота"
}

def ButtonsForDaysOfWeek():
  markup = types.InlineKeyboardMarkup(row_width=2)
  mnBt = types.InlineKeyboardButton(list(dictForDays.values())[0], callback_data=list(dictForDays.keys())[0])
  tuBt = types.InlineKeyboardButton(list(dictForDays.values())[1], callback_data=list(dictForDays.keys())[1])
  wdBt = types.InlineKeyboardButton(list(dictForDays.values())[2], callback_data=list(dictForDays.keys())[2])
  thBt = types.InlineKeyboardButton(list(dictForDays.values())[3], callback_data=list(dictForDays.keys())[3])
  frBt = types.InlineKeyboardButton(list(dictForDays.values())[4], callback_data=list(dictForDays.keys())[4])
  stBt = types.InlineKeyboardButton(list(dictForDays.values())[5], callback_data=list(dictForDays.keys())[5])
  markup.add(mnBt, tuBt, wdBt, thBt, frBt, stBt, row_width=2)
  return markup


def GetArchiveName(callbackData):
  archiveName = ""
  content = os.listdir()
  for i in range(len(content)):
    if callbackData in content[i]:
      archiveName = content[i]
  return archiveName


def ZipFileWithScheduleOpen(archiveName, fileName):
  with ZipFile(archiveName, "r") as archive:
    with archive.open(fileName, "r") as scheduleFile:
      schedule = scheduleFile.read()
      schedule = schedule.decode('utf-8')
  return schedule


def Raspes(archiveName, day):
  text = ZipFileWithScheduleOpen(GetArchiveName(archiveName), f"{day}.txt")
  return str(text)


def NamesForButtons():
  groupNames = list()
  content = os.listdir()
  for i in range(len(content)):
    if ".zip" in content[i]:
      groupNames.append(content[i][:3])
  return groupNames


def ButtonsCreate():
  groupNames = NamesForButtons()
  buttonNames = list()
  for i in range(len(groupNames)):
    buttonNames.append(
      types.InlineKeyboardButton(groupNames[i], callback_data=groupNames[i]))
  return buttonNames


def DeleteButtonsWithGroupNames(message, groupName):
  bot.edit_message_text(chat_id=message.chat.id, message_id=message.message_id, text=f"Вы выбрали группу: {groupName}", reply_markup=None)


def DeleteButtonsWithDays(message):
  bot.edit_message_text(chat_id=message.chat.id, message_id=message.message_id, text="День выбран", reply_markup=None)


def welcome(message):
  buttons = ButtonsCreate()
  markup = types.InlineKeyboardMarkup(row_width=6)
  markup.add(*buttons, row_width=6)
  bot.send_message(message.chat.id, text="Выберите группу:", reply_markup=markup)


APIKey = '5694283534:AAE3wFqbzNx2StopANXwcKsx214H3Zp-gcs'
bot = telebot.TeleBot(APIKey)


@bot.message_handler(commands=['start'])
def start(message):
  try:
    if message.chat.id not in [i[0] for i in db.returnAllUsers()]:
      db.addInfoToDB(user_id=message.chat.id, user_name=message.chat.username)
  except Exception as e:
    for id in db.returnAdmins():
      bot.send_message(id, f"Ошибочка:\n{e}")
  finally:
    print(db.returnTable())
    welcome(message)


@bot.callback_query_handler(func=lambda call: True)
def callback_inline(call):
  try:
    if call.data in list(dictForDays.keys()):
      bot.send_message(
        call.message.chat.id,
        text=f"Расписание на {dictForDays[call.data]}: \n" +
        Raspes(dictForCallbackData[call.message.chat.id], call.data))
      dictForCallbackData.pop(call.message.chat.id)
      welcome(call.message)
      DeleteButtonsWithDays(call.message)
    elif call.data in NamesForButtons():
      dictForCallbackData[call.message.chat.id] = call.data
      DeleteButtonsWithGroupNames(call.message, dictForCallbackData[call.message.chat.id])
      dOW = ButtonsForDaysOfWeek()
      bot.send_message(call.message.chat.id, text="Выберите день", reply_markup=dOW)
  except Exception as e:
    for i in [i[0] for i in db.returnSuperUsers()]:
      bot.send_message(i, f"косяк:\n{e}")
  finally:
    if call.message.chat.id not in [i[0] for i in db.returnAllUsers()]:
      bot.send_message(
        call.message.chat.id, "В боте вышло обновление, нажмите /start чтобы изменения вступили в силу")


@bot.message_handler(commands=['set_admin'])
def reqUserIdToSet(message):
  if message.chat.id in [i[0] for i in db.returnSuperUsers()]:
    message_with_request_user_id = bot.send_message(
      message.chat.id,
      "Введите user_id пользователя, чтобы сделать его администратором")
    bot.register_next_step_handler(message=message_with_request_user_id, callback=setAdmin)
  else:
    bot.send_message(message.chat.id, text="Пошел на хуй, чурка")


def setAdmin(message):
  try:
    if message.text:
      db.setToAdminStatus(user_id=int(message.text))
      bot.send_message(message.chat.id, "Выполнено")
    else:
      raise ValueError
  except ValueError:
    bot.send_message(message.chat.id, "user_id - число, дурень")


@bot.message_handler(commands=['delete_admin'])
def reqUserIdToDel(message):
  if message.chat.id in [i[0] for i in db.returnSuperUsers()]:
    message_with_request_user_id = bot.send_message(
      message.chat.id,
      "Введите user_id пользователя, чтобы забрать его права администратором")
    bot.register_next_step_handler(message=message_with_request_user_id, callback=deleteAdmin)
  else:
    bot.send_message(message.chat.id, text="У вас недостаточно прав")


def deleteAdmin(message):
  try:
    if message.text:
      db.deleteAdminStatus(user_id=int(message.text))
      bot.send_message(message.chat.id, "Выполнено")
    else:
      raise ValueError
  except ValueError:
    bot.send_message(message.chat.id, "user_id - число, дурень")


@bot.message_handler(commands=['give_me_my_user_id_please'])
def give_me_my_user_id_please(message):
  if message.chat.id not in [i[0] for i in db.returnAllUsers()]:
    db.addInfoToDB(message.chat.id, message.chat.username)
  bot.send_message(message.chat.id, f"Твой user_id: {message.chat.id}")


@bot.message_handler(content_types=['document'])
def handle_docs(message):
  if message.chat.id in [i[0] for i in db.returnAdmins()]:
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    with open(message.document.file_name, "wb") as new_file:
      new_file.write(downloaded_file)
    ShedRed.ScheduleCreate()
    bot.send_message(message.chat.id, text="Okay... I download it.")
    bot.send_message(message.chat.id, text="Введите /start")
  else:
    bot.send_message(message.chat.id, text="У вас недостаточно прав")


@bot.message_handler()
def PoshelNahui(message):
  bot.send_message(message.chat.id, text="Введите /start")


keep_alive()
bot.polling(none_stop=True, interval=0)
