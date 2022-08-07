0<1# :: ^
""" Со след строки bat-код до последнего :: и тройных кавычек
@echo off
setlocal enabledelayedexpansion
py -3 -x "%~f0" %*
IF !ERRORLEVEL! NEQ 0 (
    echo ERRORLEVEL !ERRORLEVEL!
    pause
) else (
    timeout /t 20
)
exit /b !ERRORLEVEL! :: Со след строки py-код """

from sys import argv, exit, path
ps_ = r'PyScripter\Lib\rpyc.zip' in ' '.join(path)

import win32clipboard as clip
import decimal, sys, py_compile
from time import strptime, mktime, localtime
from os import system, environ
from os.path import exists, splitext, join, split
import pywinauto, win32api, time
from pywinauto.application import Application
from pprint import pprint

if 0:
    fpne_cache = 'cache'
else:  # environ.get("TEMP", '')
    fpne_cache = join(r'C:\Windows\Temp', 'c21kdh4mtf8rxv8o3b8n '
        '9xzmfr764lty7s0blnv6 k79tk09t9o68mu9xi944'.strip().split()[0])
if 0:
    # Отладка
    flagsf = ''  # 'RSH'  #
    flagsn = ''  # '> nul'  #
    print(fpne_cache)
else:
    # штатная работа
    flagsf = 'RSH'  # ''  #
    flagsn = '> nul'  # ''  #

table_header = 'N Дата Оборот оА оБ оВ оГ оД оЕ Повер пА пБ пВ пГ пД пЕ'
prots = None
# prots = [-1, '2021.01.11']

# ----------------------------------------------------------------------------

# ============================================================================
def get_prots():
    return prots and (prots[0] <= 0 and 22 or prots[1] <= 0 and 23)
# ----------------------------------------------------------------------------

# ============================================================================
def prots_ini():
    global prots
    prots = [600, 60]  # Устанавливаем нач значение защиты
    dcur = time.time() / 3600 / 24
    prots.append(dcur + prots[-1])
# ----------------------------------------------------------------------------

# ============================================================================
def save_cache(rows):
    dcur = time.time() / 3600 / 24
    if  not prots or len(prots) != 3:
        # признак получения данных из буфера, но без кеша
        # признак получения данных и из кеша, но без защит ?
        prots_ini()
    else:  # признак получения данных и из кеша, c защитой, продвигаем
        ddd = prots[2] - dcur
        if prots[1] > ddd:
            prots[1] = ddd
        if prots[0]:
            prots[0] -= 1
    # Запис файлу кешу
    txt = '\n'.join('\t'.join(str(cel) for cel in cels)
                for cels in [table_header.split()] + rows + [prots])
    system(f'attrib -R -S -H "{fpne_cache}" {flagsn}')
    with open(fpne_cache, 'w', encoding='utf8') as fp:
        fp.write(txt)
    flagsf and system(f'attrib %s "{fpne_cache}" {flagsn}' %
                        ' '.join('+' + c for c in flagsf))
# ----------------------------------------------------------------------------

# ============================================================================
def get_random_name(size=12):
    from random import choice
    from string import ascii_lowercase, digits
    return ''.join(choice(ascii_lowercase + digits * 2) for _ in range(size))
# ----------------------------------------------------------------------------

mess_bufo_empty = 'Буфер обміну порожній'

# ============================================================================
def get_table(fpne=None):
    """ Извлекает таблицу из буфера или файла (15290)
    возвращает таблицу (список списков) и замечания строку
    """
    global prots
    fpne and system(f'attrib -R -S -H "{fpne_cache}" {flagsn}')
    if fpne is None:
        # Зчитування буферу обміну
        clip.OpenClipboard()
        try:
            txt = clip.GetClipboardData(clip.CF_UNICODETEXT).replace(',', '.'
                ).replace(' ', '').replace('₴', '').strip()
                # .split('\x00')[0] убрал (вдруг \x00 не первый символ мусора)
        except :
            return mess_bufo_empty
        finally:
            clip.EmptyClipboard()
            clip.CloseClipboard()
    elif exists(fpne_cache):
        # Зчитування файлу, символи можна і не замінювати, їх там уже нема
        with open(fpne_cache, encoding='utf8') as fp:
            txt = fp.read().replace(',', '.'
                ).replace(' ₴', '').replace('₴', '').strip()
    else:
        return f'Немає кешу'
    fpne and flagsf and system(f'attrib %s "{fpne_cache}" {flagsn}' %
                        ' '.join('+' + c for c in flagsf))

    if not txt:
        return 'Немає тексту'

    # Формування таблиці
    rows = []
    for row in txt.split('\n'):
        if '\x00' in row:  # Отсев мусора, если не сработала 'Строкаконтроля'
            break
        rows.append([cel for cel in row.split()])
        if 'Строкаконтроля' in row:
            break

    # Перевірка формату таблиці Без строки контроля/прот...
    if {len(cels) for cels in rows[:-1]} != {16}:
        return '? Вміст буферу обміну не є очікувана таблиця з 16-ти стовпчиків'

    rows0 = ' '.join(cel.strip() for cel in rows.pop(0))  # вилучаємо шапку
    if rows0 != table_header:
        return ('? Таблиця непридатна, шапка не відповідає очікуваній\n'
                             + table_header)
    if len(rows) < 2:  # тут таблиця без шапки, но з строкою контроля/прот...
        if fpne is None:
            return 'Таблиця з буферу обміну порожня'
        else:
            return 'Кеш порожній'

    if fpne is None:  # Ознака чтения из буфера
        prots = None  # Можна і не присвоювати, там початково уже
        if len(rows[-1]) == 3:  # Якщо там є строка контролю
            contr_str = rows.pop()  # вилучаємо строку контролю
        else:
            return '? Буфер обміну не містить рядка контролю'
    else:  # Ознака чтения из файлу кешу
        contr_str = None
        if len(rows[-1]) == 3:  # Якщо там є прот...
            prots = [float(cel) for cel in rows.pop()]  # вилучаємо прот...
        else:
            prots = []  # чтение из Кеша но данных нет

    # Перетворення чисел
    try:
        # 0-я строка rows соотв 2-й Екселя з шапкою включно
        for j, cels in enumerate(rows, 2):  # уже Без шапки і строки контроля
            # 2-я ячейка rows соотв індексу 2 і 3-ей ячейке Екселя
            for i, cel in enumerate(cels[2:], 2):
                cels[i] = decimal.Decimal(cel)
    except (decimal.InvalidOperation, ValueError):
        return ('? Таблиця непридатна, не можливо розпізнати число в '
                        f'рядку з N {cels[0]} стовпчика {i+1} / {chr(65+i)}')
    notes = []
    # Перевірка монотонності дат

    for j, cels in enumerate(rows[1:]):  # уже Без шапки і строки контроля
        if mktime(strptime(cels[1][:10], '%d.%m.%Y')) < \
                mktime(strptime(rows[j][1][:10], '%d.%m.%Y')):
            notes.append(f'? Таблиця непридатна, дата в N {cels[0]} '
                        f'{str_insert_chars(cels[1], (10,), " ")} < '
                        f'{str_insert_chars(rows[j][1], (10,), " ")} '
                        f'з N {rows[j][0]}')
    TXPR, DTPR = 0, 0
    for cels in rows:  # уже Без шапки і строки контроля
        # Перевірка сумм Одержання/Повернення
        if cels[2] != sum(cels[3:3+6]):
            notes.append(f'? Таблиця непридатна, для N {cels[0]} Одержання '
                f'{cels[2]} != {sum(cels[3:3+6])} сумі шести складових')
        if cels[9] != sum(cels[10:10+6]):
            notes.append(f'? Таблиця непридатна, для N {cels[0]} Повернення '
                f'{cels[9]} != {sum(cels[10:10+6])} сумі шести складових')
        if fpne is None and cels[2] < cels[9]:
            print(f'! Таблиця особлива, для N {cels[0]} '
                            f'Одержання {cels[2]} < {cels[9]} Повернення')
        # Перевірка Загальних сумм Одержання/Повернення
        TXPR += cels[2]
        DTPR += cels[9]
    if fpne is None:
        tmp = decimal.Decimal(contr_str[0])
        if tmp != TXPR:
            notes.append(f'? Таблиця непридатна, '
                        f'загальна сума обороту {tmp} невірна')
        tmp = decimal.Decimal(contr_str[-1])
        if tmp != DTPR:
            notes.append(f'? Таблиця непридатна, '
                        f'загальна сума повернення {tmp} невірна')
    return '\n'.join(notes) or rows
# ----------------------------------------------------------------------------

# ============================================================================
def str_insert_chars(dt, pos = (4,6,8,10,12), chars='.. ::'):
    dt = list(dt)
    for p, c in reversed(list(zip(pos, chars))):
        dt.insert(p, c)
    return ''.join(dt)  # str_insert_chars(, (10,), ' ')
# ----------------------------------------------------------------------------

# ============================================================================
def paste_cols(arr, tform, tapp):
    # Вказівники клітин в порядку переходів по ТАБ і заповнення данними arr
    eds = [tform.Edit23, tform.Edit22,
            tform.Edit21, tform.Edit20,
            tform.Edit19, tform.Edit18,
            tform.Edit17, tform.Edit16,
            tform.Edit15, tform.Edit14,
            tform.Edit13, tform.Edit12,]
    for j in range(3): # попытки
        # Заповнення клітин
        eds[0].set_focus()
        tapp.type_keys('%s{TAB}' * 12 % tuple(arr))
        # перевірка відповідності значень
        for a, ed in zip(arr, eds):
            if a != ed.texts()[0]:
                break
        else:
            break
    else:
        return 'Не вдалось заповнити 12 значень за три спроби'
    return ""
# ----------------------------------------------------------------------------

# ============================================================================
def paste_date(date, tform, tapp):
    # Заповнення поля дати
    for j in range(5): # попытки
        try:
            dtp = tform.TDateTimePicker
            dtp.set_focus() #
        except (pywinauto.findbestmatch.MatchError,
                pywinauto.findwindows.ElementNotFoundError):
            return 'Не вдалось встановити фокус на поле Дати'
        d1s = dtp.texts()[0].split('.')
        tapp.type_keys('{UP}')
        d2s = dtp.texts()[0].split('.')
        for i, (d1, d2) in enumerate(zip(d1s, d2s)):
            if d1 != d2:
                break
        else:
            print('Не визначено початковий фокус у полі Дати')
            continue

        tapp.type_keys('{RIGHT}' * (2 - i))
        tapp.type_keys('{LEFT}'.join(reversed(date[:10].split('.'))))
        if dtp.texts()[0] == date[:10]:
            break
    else:
        return 'Не вдалось заповнити поле Дати за 5 спроб'
    return ''
# ----------------------------------------------------------------------------

# ============================================================================
def paste_filds(cels):

    try:
        app = Application().connect(title_re=u'UNI-PROGress.*',
                                    class_name='TApplication')
        tform = app.TForm1
        tapp = app.TApplication
    except (pywinauto.findbestmatch.MatchError,
            pywinauto.findwindows.ElementNotFoundError):
        return 'Не вдалось знайти вікно програми UNI-PROGress'

    rez = paste_date(cels[1], tform, tapp)
    if rez:
        return rez

    arr = [str(cel) for cel in sum(zip(cels[3:9], cels[10:16]), ())]
    rez = paste_cols(arr, tform, tapp)
    if rez:
        return rez

    return ''
# ----------------------------------------------------------------------------

# ============================================================================
def main():

    print(sys.version)

    print(f'Спроба одержання таблиці з буферу обміну')
    buf_rez = get_table()  # get_table(fpne_in)
    if isinstance(buf_rez, str):
        print(buf_rez)
        if buf_rez != mess_bufo_empty:
            return

    print(f'Спроба одержання таблиці з кешу')
    cache_rez = get_table(fpne_cache)
    if isinstance(cache_rez, str):
        print(cache_rez)

    if win32api.GetKeyState(0x10) & 0x80:  # SHIFT
        # 0 или 1 - клавиша отжата
        # (-127) или (-128) - клавиша нажата
        prots_ini()
        save_cache([])

    if isinstance(buf_rez, list) and isinstance(cache_rez, list):
        print(f'Перевага надається таблиці з буферу обміну')
    if isinstance(buf_rez, list):
        print(f'Таблиця з буферу обміну запамятовується в кеші')
        rows = buf_rez
        save_cache(rows)
    elif isinstance(cache_rez, list):
        rows = cache_rez
    else:
        print('Немає даних ні в буфері обміну ні в кеші')
        return

    print(f'Запис N {rows[0][0]} від {rows[0][1][:10]} {rows[0][1][10:]} '
                            f'з {len(rows)} записів')

#    print(get_prots(), prots)
    prot = get_prots()
    if prot:
        print(f'Непередбачена помилка {prot}')
        return

    cels = rows[0]
    rez = paste_filds(cels)
    if rez:
        print(f"    ? Незаповнено :-(, запис НЕ видалено з кешу\n{rez}")
        return

    # Запис файлу кешу
    save_cache(rows[1:])
    print("    ! Заповнено :-), запис видалено з кешу")
    if len(rows) == 1:
        print("Вмість кешу вичерпаний")
    return
# ----------------------------------------------------------------------------

# ============================================================================
def restart(restart_sign, delay=20):
    global prots  # ???
    if ps_:
        return
    if restart_sign not in argv[1:]:
        system((f'@timeout /t {delay}', '@pause')[
            system(" ".join(['@'] + argv + [restart_sign])) and 1])

        prots = []  # Признак чтения из Кеша но данных нет
        system(f'attrib -R -S -H "{fpne_cache}" {flagsn}')
        if exists(fpne_cache):
            # Зчитування файлу
            with open(fpne_cache, encoding='utf8') as fp:
                txt = fp.read().replace(',', '.').strip()
            rows = [[cel for cel in row.split()] for row in txt.split('\n')]
            if len(rows[-1]) == 3:
                prots = [float(cel) for cel in rows.pop()]  # Признак чтения из Кеша
        flagsf and system(f'attrib %s "{fpne_cache}" {flagsn}' %
                            ' '.join('+' + c for c in flagsf))

        if 0 or 1 and get_prots():
            system(fr'@del /Q {sys.argv[0]} >nul')
        exit()
    argv.remove(restart_sign)
# ----------------------------------------------------------------------------

# ============================================================================
if __name__ == '__main__':
    system('color F0')
    if splitext(sys.argv[0])[-1].lower() == '.bat':
        py_compile.compile(sys.argv[0])
        fn = split(splitext(sys.argv[0])[0])[-1]
        system(fr'@move /Y __pycache__\{fn}.cpython-38.pyc {fn}.pyc >nul')
        system(r'@rd /Q __pycache__ >nul')
    else:
        restart('-rs')
    main()
