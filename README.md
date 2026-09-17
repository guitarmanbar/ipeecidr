# ipeecidr
Tool for local IP calculating

Как пользоваться.

Установите зависимость: pip install openpyxl

Запустите скрипт: python main.py

Команды в REPL:

1. load network.xlsx (инициализирует модель подсетей);
2. lookup D:\projects\ipeecidr\ips.txt D:\projects\ipeecidr\network.xlsx 
output.xlsx (позволяет по списку IP-адресов из файла ips.xlsx или ips.txt получить информацию о подсетях, куда эти IP-адреса входят). Примеры файлов:
network.xlsx. URL: https://github.com/user-attachments/files/32353016/network.xlsx
output.xlsx. URL: https://github.com/user-attachments/files/32353017/output.xlsx
3. aggregate D:\projects\ipeecidr\ips.txt D:\projects\ipeecidr\output2.xlsx (склеивает несколько IP-адресов в одну общую подсеть; если подсетей несколько - выдаёт несколько подсетей; на данный момент работает некорректно). Примеры файлов:
output2.xlsx. URL: https://github.com/user-attachments/files/32353020/output2.xlsx
4. exit (выход из REPL)

Особенности реализации.

1. Модель сети загружается из Excel, колонки определяются по вхождению подстрок cidr, name, desc.
2. Входные IP читаются либо из Excel (первая колонка), либо из текстового файла (по одному в строке). Тип определяется по расширению.
3. Дубликаты удаляются, пробелы обрезаются.
4. Невалидные IP вызывают исключение с трассировкой.
5. Функция 1: для каждого IP выводит все совпавшие подсети; если совпадений несколько, ставит флаг Duplicate = True; если нет — NOT_FOUND.
6. Функция 2: сортирует IP, группирует в непрерывные диапазоны и с помощью ipaddress.summarize_address_range получает минимальный набор CIDR, точно покрывающий только эти адреса. Результат сохраняется в Excel с колонками CIDR, маска, количество IP.
7. Вывод обеих функций сохраняется в Excel.
