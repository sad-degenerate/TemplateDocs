# TemplateDocs / Библиотека для создания документов по шаблону

## Как пользоваться?

В этой библтотеке есть 2 класса, которые отвечают за 2 функции: замена по шаблону и печать.

### Замена по шаблону

Для замены слов в документе по шаблону вам нужно создать новый экземпляр класса "DocumentReplacer", передав в него путь к шаблону документа, по которому будет производится замена, а так же путь к папке, в которую будут отправлятся результаты работы программы.
Далее нужно вызвать метод "Replace", в который передать словарь из пар слов [слово, которое нужно заменить | слово, на которое нужно заменить] и имя файла с результатами.
После этого произойдет сама замена и будет создан новый документ по шаблону.

### Печать

Для того чтобы распечатать любой Word документ достаточно создать новый экземпляр класса "DocumentPrinter", передав в конструктор путь к файлу, который вы хотите распечатать.
После этого нужно вызвать метод класса "Print", передав туда число копий, и ваши документы распечатаются на принтере, установленном по умолчанию.
