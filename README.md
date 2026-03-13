# CHECK_INN

консольна програма для перевірки файлів, в яких використовуються РНОКПП (Україна). 
Знаходить 10-значні цифри і перевіряє на валідність. Виводить повідомлення про результат.

Використання:
в терміналі (cmd, PowerShell) запускаємо \шлях\до\check_inn.exe \шлях\до\файл.{odt,doc,docx,rtf}


![check_inn](Screenshot%202025%2D11%2D28%2016.40.17.png)

можна використовувати для пакетної обробки файлів

Linux

```
find . -type f -name "*.txt" -exec check_inn {} \;for file in *.txt; do
```

Windows
Создать run_all.bat
```
@echo off
for %%f in (*.txt) do (
    check_inn %%f
)
```
