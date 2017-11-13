Some bash-like python skripts for mining data from Sketch Engine and import into xlsx files.  
Dockalk pull data from sketch engine corpus, calc some fitch, and impori it into xlsx file.  
StaffParse take xlsx file, parse it, calculate some fithes and write it into another xlsx file.  
<jf>
Скрипты реализованы на Python 3, поэтому используют bash - подобный синтаксис. Для запуска, помимо интерпретатора Python(для windows можно скачать здесь https://www.python.org/downloads/windows/) необходимы следующие внешние библиотеки: openpyxl(устанавливается с помощью команды pip3 install openpyxl) и scipy(pip3 install scipy)
<jf>  
<jf>
**docalk.py** - Скрипт для получения данных и подсчёта корреляций между метриками.  
Ключи:   
**-q** - обязательный параметр, используется для поиска определённой леммы/словосочетания/и.т.д. Используется cql-синтаксис.  
**Примечания: Python - скрипты игнорируют спец. символы, поэтому при вводе запроса их необходимо экранировать.** Т.е. Запрос вида: [lemma="ночь"] должен выглядить как [lemma=\"ночь\"], иначе скрипт выдаст ошибку.  
**-n** - имя пользователя Sketch Engine  
**-k** - ключ для доступа к api Sketch Engine, получить можно здесь: https://the.sketchengine.co.uk/auth/api_access/  
**-cor** - необязательный параметр. Корпус с которым будет работать скрипт(по умолчанию utenten11_8_1G)  
**-mit** - необязательный параметр. Колличество коллокаций, которые необходимо проанализировать(по умолчанию 300)  
**-f** - необязательный параметр. Нижнее значение окна анализа коллокаций(по умолчанию -3)  
**-t** - необязательный параметр. Верхнее значение окна анализа коллокаций(по умолчанию 1)  
Простейшая команда для запуска скрипта выглядит так:
```python
python3 docalk.py -q [lemma=\"ночь\"] -n ваше_имя_пользователя -k ваш_ключ_доступа
```
Скрипт сгенерирует отчёт и сохранит его в файле results.xlsx(отсчёт будет лежать в той же папки, откуда происходил запуск скрипта)  
<jf>
<jf>
**staffParse.py** - Скрипт для подсчёта среднего, оптимизированного и нормированного рангов, а так же количества мер.  
**Внимание: для скрипта критически важно, чтобы значения мер находились в столбцах с С по L, иначе это приведёт к неправильному подсчёту рангов и другим ошибкам. Так же, важно чтобы значения мер находились в .xlsx файле на первом листе, иначе скрипт работать не будет.**  
Ключи:  
**-n** - обязательный параметр. Имя файла, с которым будет работать скрипт  
**-qv** - необязательный параметр. Количество коллокаций, для которых будет высчитыватся столбец "количество мер". Для всех других коллокаций колличество мер бует считаться равным 0.  
**-out** - необязательный параметр. Имя файла, в котором сохранится результат раоты скрипта(по умолчанию result.xlsx)  
Простейшая команда для запуска будет выглядить так:
```python 
python3 staffParse.py -n имя_файла_с_метриками
```
Скрипт сгенерирует отчёт и сохранит их в файл result.xlsx, если не указан ключ **-out**
<jf>