# DocTempl

## Формирование документов из шаблонов MS Word

Данный скрипт предназначен для форммирования пакета документов на основе заранее созданных *.docx шаблонов, с заполнением этих шаблонов из файла с данными для подстановки

## Установка

Скачать

## Пример запуска

В каталоге 2021.01 выполнить команду 

    powershell.exe -ExecutionPolicy Bypass -File ..\DocTempl.ps1 РогаИКопыта-Иванов.txt

После выполнения скрипта в нём будут созданы три файла

    ИП Иванов И.В._ФабрикаАБВГД_Январь 2021 г._Акт сдачи-приемки оказанных услуг № 777.docx  
    ИП Иванов И.В._ФабрикаАБВГД_Январь 2021 г._Заявка на услуги № Ф-123.docx  
    ИП Иванов И.В._ФабрикаАБВГД_Январь 2021 г._Отчёт об оказанных услугах № Ф-123.docx  

## Описание файла данных

В каталоге 2021.01 находится пример [файла](2021.01/РогаИКопыта-Иванов.txt)

Файл данных состоит из двух секций 

### Секция [Шаблоны] 

Cодержит список обрабатываемых шаблонов, в каждой строке секции указывается путь и имя файла шаблона

### Секция [Поля] 

Cодержит список полей с данными, которые будут в поля шаблонов из первой секции
В этой секции имена переменных отделены от значения символом **=**

## Содание шаблонов

Примеры шаблонов расположены в каталоге [template](template)

В docx файлах в необходимые места вставляются поля с типом "DocVariable", которым задаются имена.
Скрипт выберет из файла данных имена переменных и заполнит их значения в каждом указанном шаблоне. 

Вставку можно выполнить с помощью пункта меню  
![](images/img3.png)  


Переключение между режимом отображения кодов полей и их представлением выполняется с помощью клавиш **Alt+F9**

Режим промсотра кодов полей  
![](images/img4.png)  

Режим представления значений полей (нормальный режим)  
![](images/img5.png)  

Пример изменения/добавления поля  
![](images/img2.png)  
![](images/img1.png)  
