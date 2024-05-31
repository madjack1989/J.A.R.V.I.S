# Программа для работы Jarvis

import pystray
import PIL 
import PIL.Image
import keyboard 
import webbrowser
import subprocess
import pyperclip as p 
import configparser  
import os
import ctypes
import win32api
import win32gui
from customtkinter import *


# Тема системы: System, light, dark
set_appearance_mode("dark")
# Темы окон: blue(defoult), dark-blue, green
set_default_color_theme("dark-blue")

#---Браузер для работы-------------------------------------------------
config = configparser.ConfigParser()            # создаём объекта парсера
config.read("config.ini",encoding='utf-8')      # читаем конфиг
b = config["browser"]["b"]                      # передаем инфу с файла conf.ini
b = b.replace("\"",'')
#----------------------------------------------------------------------

#---Автостарт софта----------------------------------------------------
config = configparser.ConfigParser()
config.read("config.ini", encoding="utf-8")
if config['check_box']['c'] == "On":
    s = config['soft']['s']
    s = s.split(",")
    for i in s:
        if i == '':
            continue
        elif i.find("Discord") != -1:
            os.system(i+" --processStart Discord.exe")
        os.startfile(i)
#----------------------------------------------------------------------


#---Функции для семны раскладки клавиатуры-----------------------------
def setCyrillicLayout():    #Установка русской раскладки
    window_handle = win32gui.GetForegroundWindow()
    result = win32api.SendMessage(window_handle, 0x0050, 0, 0x04190419)
    return(result)

def setEngLayout():         #Установка английской раскладки
    window_handle = win32gui.GetForegroundWindow()
    result = win32api.SendMessage(window_handle, 0x0050, 0, 0x04090409)
    return(result)
      
def get_layout():           #Получение текущей раскладки
    u = ctypes.windll.LoadLibrary("user32.dll")
    pf = getattr(u, "GetKeyboardLayout")
    if hex(pf(0)) == '0x4190419':
        return 'ru'
    if hex(pf(0)) == '0x4090409':
        return 'en'
#----------------------------------------------------------------------

# Удаление пробелов
def ip_trim(ip):
    ip = ip.replace(' ', '')
    return ip

# Удаление /порт
def del_port(ip):
    ip = ip.split('/')
    del ip[1]  
    return " ".join(ip)

# Обработка запроса GPON ветки
def branch_trim(branch):
    branch = branch.replace(' ', '')

# Возвращает только IP 
def just_ip(ip):
    ip = ip_trim(ip)
    if ip.find('/') > 0:
        ip = del_port(ip)
    return ip  

# Удаление /порт
def del_pon(ip):
    ip = ip.split(':')
    del ip[0]  
    return "".join(ip)

# Удаление лишнего GPON
def gpon_trim(pon):
    pon = ip_trim(pon)
    if pon.find('GPON') == 0:
        pon = del_pon(pon)
    if pon.count('/') == 4:
        pon = pon[0:pon.rfind('/')]
    return pon

# Формирование запроса для ge порта
def trim_signal_ge(var):
    var = ip_trim(var) # Удаляем пробелы по сторонам
    var = var.split('/') # Разделяем
    port = var[1].split(':') # Собираем порт
    port = '%2F0%2F'.join(port)
    var='ge-'+port # Собираем всю конструкцию
    return var

# Формирование запроса для xe порта
def trim_signal_xe(var):
    var = ip_trim(var) # Удаляем пробелы по сторонам
    var = var.split('/') # Разделяем
    port = var[1].split(':') # Собираем порт
    port = '%2F0%2F'.join(port)
    var='xe-'+port # Собираем всю конструкцию
    return var

# Телнет
def telnet():
    ip = p.paste()                                                                      # берет из буфера скопированный IP 
    ip = just_ip(ip)                                                                    # Удаляем все лишнее
    subprocess.Popen(args=['telnet', ip], creationflags=subprocess.CREATE_NEW_CONSOLE)  # Откывает telnet в отдельном окне с переданным IP 

# Веб морда   
def web_m():
    ip=p.paste()                                                                         
    ip = just_ip(ip)                                                                        
    webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(b))               # Определение вторичного браузера путь к которому берем из conf.ini
    webbrowser.get('firefox').open(ip)                                                  # Открываем вкладку   
    
# Пинг
def ping():
    ip=p.paste()                                                                            
    ip = just_ip(ip)
    subprocess.Popen(args=['ping', '-t', ip], creationflags=subprocess.CREATE_NEW_CONSOLE)  # Откывает ping в отдельном окне с переданным IP

# Кто прибинден
def bind():
    ip=p.paste()                                                                           
    ip = just_ip(ip)
    url = 'https://billing.briz.ua/Ru/billing/user/finder.html?filter=inet&sw='+ip+'&sb=3'  # Формируем ссылку с IP
    webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(b))                   
    webbrowser.get('firefox').open(url)
                                                          
# Диагностика GPON по ветке
def gpon():
    branch = p.paste()                                                                                  # берет из буфера ветку
    branch = gpon_trim(branch)
    url = 'https://billing.briz.ua/Ru/admin/gpon/gpon_onu-check.html?userId=&gponName%5B%5D='+branch    # Формирует запрос + ветка
    webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(b))                               
    webbrowser.get('firefox').open(url)                                                                

# Дерево устройств
def tree():
    ip=p.paste()                                                                            
    ip = just_ip(ip)
    url = 'http://sip.briz.ua/switch_tree?ip='+ip                             # Формируем ссылку с IP
    webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(b))                   
    webbrowser.get('firefox').open(url) 
    
# Порты джуна    
def ports():
    var = p.paste()
    var = just_ip(var)
    url = 'http://juny.briz.ua/juniper/ports?ip='+var
    webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(b))                   
    webbrowser.get('firefox').open(url)              

# Сигналы на дом
def signal():
    var = p.paste()
    ip = ip_trim(var)
    ip = del_port(ip)
    port1 = trim_signal_ge(var)
    port2 = trim_signal_xe(var)
    url1 = 'http://juny.briz.ua/juniper/signals?csrfmiddlewaretoken=9WybtPB2zkKWd98MNBhHPnyq2hmxhtLDsCNLHI5WtbKmDPivEnjeIQGtSoCTsya9&ip='+ip+'&port_type=physical&port='+port1
    url2 = 'http://juny.briz.ua/juniper/signals?csrfmiddlewaretoken=9WybtPB2zkKWd98MNBhHPnyq2hmxhtLDsCNLHI5WtbKmDPivEnjeIQGtSoCTsya9&ip='+ip+'&port_type=physical&port='+port2
    webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(b))                   
    webbrowser.get('firefox').open(url1)
    webbrowser.get('firefox').open(url2)
    
# Оборудование по IP
def equipment():
    ip=p.paste()                                                                           
    ip = just_ip(ip)
    url = 'https://billing.briz.ua/Ru/admin/equipment.html?ip='+ip  # Формируем ссылку с IP
    webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(b))                   
    webbrowser.get('firefox').open(url)

# Вланы
def vlan():
    url = 'https://billing.briz.ua/Ru/reports/engineers/overloadSubnets.html'  # Формируем ссылку с IP
    webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(b))                   
    webbrowser.get('firefox').open(url)

# Автологин в telnet
def auto_login():
    if get_layout() == 'ru':
        setEngLayout()
    keyboard.write('\b\badmin')
    keyboard.press_and_release('Enter')
    keyboard.write('Masterok')
    keyboard.press_and_release('Enter')


#---Запуск программы в трее----------------------------
image = PIL.Image.open('icon.png')  

# Запуск окна настроек программы
def open_config_window():
    
    #---Окно настроек---------------------------------------------------------------------------------------------------
    #---Browser---
    # Вводит данные в поле browser
    def browser_path_input():
        config = configparser.ConfigParser()
        config.read("config.ini", encoding="utf-8")
        browser = config['browser']['b']
        browser = browser.replace("\"",'')
        path.insert(0, browser)
        
    # Действия по нажатию кнопки Open (Открывает диалогововое окно и перезаписывает путь в поле браузера браузера)
    def chenge_path_brawser():
        window.filename = filedialog.askopenfilename(initialdir="C:/Users/JackOS/Desktop", title="Какой файл открываем", filetypes=(("exe файлы",".exe"),("png файлы", "*.png"),("Все файлы","*.*")))  
        END = len(path.get())
        path.delete(0, END)
        path.insert(0, str(window.filename))  
        config = configparser.ConfigParser()
        config.read("config.ini", encoding="utf-8")
        config['browser']['b']=str(window.filename)
        with open('config.ini', 'w', encoding="utf-8") as configfile:    
            config.write(configfile) 
        path.delete(0, END)    
        browser_path_input()     
            
    # Действия по нажатию кнопки Тест (Открывает www.google.com в браузере путь к которому указан в config.ini)
    def test_open_browser():
        config = configparser.ConfigParser()
        config.read("config.ini", encoding="utf-8")
        browser = config['browser']['b']
        browser = browser.replace("\"",'')
        webbrowser.register('Browser', None, webbrowser.BackgroundBrowser(browser))
        webbrowser.get('Browser').open('http://www.google.com')  
    #------

    #---Soft auto start---
    # Очистить frame
    def clear_frame():
        for widgets in frame_autostart.winfo_children():
            widgets.destroy()
        
    # Добавить новый путь к софт
    def add_soft_path():
        window.filename = filedialog.askopenfilename(initialdir="C:/", title="Какой файл открываем", filetypes=(("exe файлы",".exe"),("png файлы", "*.png"),("Все файлы","*.*")))
        config = configparser.ConfigParser()
        config.read("config.ini", encoding="utf-8")
        old_path = config['soft']['s']     
        new_path = old_path + "," + str(window.filename)
        config['soft']['s'] = new_path
        with open('config.ini', 'w', encoding="utf-8") as configfile:    
            config.write(configfile)
        draw_soft_path() 
        
    # Удалить последний путь к софт
    def del_soft_path():
        config = configparser.ConfigParser()
        config.read("config.ini", encoding="utf-8")
        old_path = config['soft']['s']
        old_path = old_path.split(",")
        new_path = old_path[:-1]
        new_path = ",".join(new_path)
        config['soft']['s'] = new_path
        with open('config.ini', 'w', encoding="utf-8") as configfile:    
            config.write(configfile)  
        clear_frame()         
        draw_soft_path()

    # Запуск софта по путям в файле config.ini        
    def test_soft_path():
        config = configparser.ConfigParser()
        config.read("config.ini", encoding="utf-8")
        old_path = config['soft']['s']
        old_path = old_path.split(",")
        for i in old_path:
            if i == '':
                continue
            elif i.find("Discord") != -1:
                os.system(i+" --processStart Discord.exe")
            os.startfile(i)
                    
    # Рисуем пути софту
    def draw_soft_path():
        clear_frame()
        
        # Перезапись чекбокса
        def res_result_check():
            config = configparser.ConfigParser()
            config.read("config.ini", encoding="utf-8")
            config['check_box']['c'] = On_Off.get()
            with open('config.ini', 'w', encoding="utf-8") as configfile:    
                config.write(configfile)
                
        config = configparser.ConfigParser()
        config.read("config.ini", encoding="utf-8")
        check = config['check_box']['c']
        rows = config['soft']['s']
        rows = rows.split(",")    
        r=1
        for i in rows:
            if i == '':
                continue
            p = i
            p = p.split("/")
            p = str(p[-1:])
            n_p = ''.join(p)
            n_p = n_p.replace(".exe","")
            n_p = n_p.replace("']","")
            n_p = n_p.replace("['","")
            n_p = n_p.upper()
            if n_p.find("UPDATE") != -1:
                n_p = "DISCORD"
            # Лейбл
            auto_start_label = CTkLabel(frame_autostart, text=str(n_p), text_color="#FFC107", font=("Comic Sans MS", 16))
            auto_start_label.grid(column=0, row=r, padx=50, pady=10)
            # Поле ввода
            auto_start_path = CTkEntry (frame_autostart, width=700, border_color="#FFC107", font=("Comic Sans MS", 16))
            auto_start_path.insert(0, i)
            auto_start_path.grid(column=1, row=r, padx=10, pady=10)
            r+=1   
        # Кнопка TEST
        test_start_button = CTkButton(frame_autostart, width=15, text='TEST', hover_color="#FFC107", fg_color="#B71C1C", text_color="#FFC107", command=test_soft_path, font=("Comic Sans MS", 16))
        test_start_button.grid(column=0, row=0, padx=10, pady=10)
        # Чекбокс
        check_status = StringVar()
        On_Off = CTkCheckBox(frame_autostart, text="Запускать софт при старте J.A.R.V.I.S", text_color="#FFC107", border_color="#FFC107", hover_color="#FFC107", fg_color="#B71C1C", variable=check_status, onvalue="On", offvalue="Off", command=res_result_check, font=("Comic Sans MS", 16))
        On_Off.grid(column=1, row=0, padx=10, pady=10) 
        # Рисуем статус взятый из config.ini
        if check == "Off":
            On_Off.deselect()
        elif check == "On":
            On_Off.select()         
        # Кнопка +
        add_soft_button = CTkButton(frame_autostart, width=50, text='+', hover_color="#FFC107", fg_color="#B71C1C", text_color="#FFC107", command=add_soft_path, font=('Helvetica',30))
        add_soft_button.grid(column=1, row=r+1, padx=250, pady=10, sticky=W)  
        # Кнопка -
        del_soft_button = CTkButton(frame_autostart, width=50, text='-', hover_color="#FFC107", fg_color="#B71C1C", text_color="#FFC107", command=del_soft_path, font=('Helvetica',30))
        del_soft_button.grid(column=1, row=r+1, padx=350, pady=10, sticky=E)    
                
        
    # Рисуем коно настроек
    window = CTk()
    window.title('Настройки J.A.R.V.I.S')
    window.iconbitmap('icon.ico')

    # Рисуем окно посередине экрана
    win_w = 1000
    win_h = 800
    screen_w = window.winfo_screenwidth()
    screen_h = window.winfo_screenheight()
    x = (screen_w / 2) - (win_w / 2)
    y = (screen_h / 2) - (win_h / 2)
    window.geometry(f'{win_w}x{win_h}+{int(x)}+{int(y)}')

    #---Браузер--------
    # Фрейм для браузера для работы
    frame_browser = CTkFrame(window)
    # Лейбл
    browser_label = CTkLabel(frame_browser, text="Браузер для работы", text_color="#FFC107", font=("Comic Sans MS", 16))
    browser_label.grid(column=0, row=0, padx=10, pady=10)
    # Поле ввода
    path = CTkEntry(frame_browser, width=600, border_color="#FFC107", font=("Comic Sans MS", 16))
    browser_path_input()
    path.grid(column=1, row=0, padx=10, pady=10)
    # Кнопка открыть
    button_open = CTkButton(frame_browser, width=15, text='Изменить', hover_color="#FFC107", fg_color="#B71C1C", text_color="#FFC107", command=chenge_path_brawser, font=("Comic Sans MS", 16))
    button_open.grid(column=2, row=0, padx=10, pady=10)
    # Кнопка TEST
    button_test = CTkButton(frame_browser, width=15, text='Тест', hover_color="#FFC107", fg_color="#B71C1C", text_color="#FFC107", command=test_open_browser, font=("Comic Sans MS", 16))
    button_test.grid(column=3, row=0, padx=10, pady=10)
    frame_browser.pack(fill = X, pady=15)
    #---Автостарт-------
    # Фрейм для автостарта программ
    frame_autostart = CTkFrame(window)
    # Рисуем пути к софту
    draw_soft_path()
    frame_autostart.pack(fill = X, pady=15)
    #-----------
    # Кнопка OK
    button = CTkButton(window, width=150, text='OK', hover_color="#FFC107", fg_color="#B71C1C", text_color="#FFC107", command=window.destroy, font=("Comic Sans MS", 16))
    button.pack(padx=10, pady=20, side='bottom') 
    window.mainloop()
#-------------------------------------------------------------------------------------------------------------

#---Окно подсказок-------------------------------------------------------------------------
def open_window_clue():
    
    def close_clue_window():
        window.destroy()
        
    
    window = CTk()
    window.title('Подсказки')
    window.iconbitmap('icon.ico')
    
    win_w = 400
    win_h = 500
    screen_w = window.winfo_screenwidth()
    screen_h = window.winfo_screenheight()
    x = (screen_w / 2) - (win_w / 2)
    y = (screen_h / 2) - (win_h / 2)
    window.geometry(f'{win_w}x{win_h}+{int(x)}+{int(y)}')

    t = "F1 - Автологин в Telnet \n Alt + F1 - Подсказки \n 1 - Telnet \n 2 - Ping \n 3 - Кто прибинден \n 4 - Втека GPON \n 5 - Веб морда \n 6 - Дерево устройств \n 7 - Порты джуна \n 8 - Сигналы на дом \n 9 - Оборудование \n 0 - Вланы"

    clue_text = CTkLabel(window, justify='left', text=t, text_color="#FFC107", font=("Comic Sans MS", 24))
    clue_text.pack(padx=20, pady=20)    
    
    clue_button = CTkButton(window, text="OK", hover_color="#FFC107", fg_color="#B71C1C", text_color="#FFC107", font=("Comic Sans MS", 16), width=100, command=close_clue_window)
    clue_button.pack(padx=20, pady=20, side='bottom')
    
    window.mainloop()
#------------------------------------------------------------------------------------------

# Действия при нажатии на пункт меню
def on_clicked(image, item):
    if str(item) == 'Настройки':
        open_config_window()
    if str(item) == '1_Телнет':
        telnet()
    if str(item) == '2_Пинг':
        ping()
    if str(item) == '3_Кто прибинден':
        bind() 
    if str(item) == '4_Ветка GPON':
        gpon()
    if str(item) == '5_Веб морда':
        web_m()
    if str(item) == '6_Дерево устройств':
        tree() 
    if str(item) == '7_Порты джуна':
        ports()    
    if str(item) == '8_Сигналы на дом':
        signal()
    if str(item) == '9_Оборудование':
        equipment()
    if str(item) == '0_Вланы':
        vlan()    
    if str(item) == 'Alt + F1 - Подсказки':
        open_window_clue()                                                         
    if str(item) == 'Выход':
        icon.stop()    
        
# Формирование меню в трее        
icon = pystray.Icon('Праграмулина!', image, title="J.A.R.V.I.S", menu=pystray.Menu(
    pystray.MenuItem('Настройки', on_clicked),
    pystray.MenuItem('1_Телнет', on_clicked),
    pystray.MenuItem('2_Пинг', on_clicked),
    pystray.MenuItem('3_Кто прибинден', on_clicked),
    pystray.MenuItem('4_Ветка GPON', on_clicked),
    pystray.MenuItem('5_Веб морда', on_clicked),
    pystray.MenuItem('6_Дерево устройств', on_clicked),
    pystray.MenuItem('7_Порты джуна', on_clicked),
    pystray.MenuItem('8_Сигналы на дом', on_clicked),
    pystray.MenuItem('9_Оборудование', on_clicked),
    pystray.MenuItem('0_Вланы', on_clicked),
    pystray.MenuItem('F1 - Автологин в telnet', auto_login),
    pystray.MenuItem('Alt + F1 - Подсказки', open_window_clue, default=True),
    pystray.MenuItem('Выход', on_clicked)
))    

# Сочитание клавиш
keyboard.add_hotkey('Alt + 1', telnet)
keyboard.add_hotkey('Alt + 2', ping)
keyboard.add_hotkey('Alt + 3', bind)
keyboard.add_hotkey('Alt + 4', gpon)
keyboard.add_hotkey('Alt + 5', web_m)
keyboard.add_hotkey('Alt + 6', tree)
keyboard.add_hotkey('Alt + 7', ports)
keyboard.add_hotkey('Alt + 8', signal)
keyboard.add_hotkey('Alt + 9', equipment)
keyboard.add_hotkey('Alt + 0', vlan)
keyboard.add_hotkey('Alt + F1', open_window_clue)
keyboard.add_hotkey('F1', auto_login)

# Запускаем окно в терее
icon.run()


# auto-py-to-exe  Прога для преобразования из py в exe