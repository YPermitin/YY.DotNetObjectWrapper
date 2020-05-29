# DotNet Object Wrapper

Оболочка объектов .NET для использования в других экосистемах приложений.

## Предыстория

Изначально проект назывался **["NetObjetToIDispatch"](https://infostart.ru/public/238584/)** и создавался силами **[Сергея Смирнова](https://infostart.ru/profile/82159/)**.

Данная разработка создавалась для использования сборок .NET в решениях на платформе 1С:Предприятие через преобразование объектов и классов в COM-объекты. Последние можно использовать в экосистеме 1С. Достигается это путем создания класса, реализующего методы интерфейса **[IReflect](https://docs.microsoft.com/ru-ru/dotnet/api/system.reflection.ireflect?view=netcore-3.1)**.

<details>
    <summary>Публикации автора по проекту</summary>

* [Использование сборок .NET в 1С 7.x b 8.x. Создание внешних Компонент](https://infostart.ru/public/238584/)
* [Обработка для формирования классов для прямого доступа к файлам 1С через курсоры BDE. И многого другого](https://infostart.ru/public/345658/)
* [Методы для группировки данных по полю,полям в Таблице Значений на примере универсального метода списания по партиям, а также отбора строк в ТЗ по произвольному условию. Для 8.x и 7.7](https://infostart.ru/public/371762/)
* [ИзСтрокиСРазделителями в Восьмерке](https://infostart.ru/public/371887/)
* [Code First и Linq to EF на примере 1С версии 7.7 и 8.3 часть I](https://infostart.ru/public/393228/)
* [Code First и Linq to EF на примере 1С версии 8.3 часть II](https://infostart.ru/public/402038/)
* [Linq to EF. Практика использования. Часть III](https://infostart.ru/public/402433/)
* [Linq to ODATA](https://infostart.ru/public/403524/)
* [.NET(C#) для 1С. Динамическая компиляция класса обертки для использования .Net событий в 1С через ДобавитьОбработчик или ОбработкаВнешнегоСобытия](https://infostart.ru/public/417830/)
* [1C Messenger для отправки сообщений, файлов и обмена данными между пользователями 1С, вэб страницы, мобильными приложениями а ля Skype, WhatsApp](https://infostart.ru/public/434771/)
* [Использование классов .Net в 1С для новичков](https://infostart.ru/public/448668/)
* [Быстрое создание Внешних Компонент на C#. Примеры использования Глобального Контекста, IAsyncEvent, IExtWndsSupport, WinForms и WPF](https://infostart.ru/public/457898/)
* [.Net в 1С. Асинхронные HTTP запросы, отправка Post нескольких файлов multipart/form-data, сжатие трафика с использованием gzip, deflate, удобный парсинг сайтов и т.д.](https://infostart.ru/public/466052/)
* [.Net в 1С. На примере использования HTTPClient, AngleSharp. Удобный парсинг сайтов с помощью библиотеки AngleSharp, в том числе с авторизацией аля JQuery с использованием CSS селекторов. Динамическая компиляция](https://infostart.ru/public/466196/)
* [Использование ТСД на WM 6 как беспроводной сканер с получением данных из 1С](https://infostart.ru/public/525806/)
* [Кроссплатформенное использование классов .Net в 1С через Native ВК. Или замена COM на Linux](https://infostart.ru/public/534901/)
* [Кроссплатформенное использование классов .Net в 1С через Native ВК. Или замена COM на Linux II](https://infostart.ru/public/541518/)
* [Асинхронное программирование в 1С через использование классов .Net из Native ВК](https://infostart.ru/public/541698/)
* [1С, Linux, Excel, Word, OpenXML, ADO, Net Core](https://infostart.ru/public/544232/)
* [.Net Core, 1C, динамическая компиляция, Scripting API](https://infostart.ru/public/547389/)
* [Net Core. Динамическая компиляция класса обертки для получения событий .Net объекта в 1С](https://infostart.ru/public/548701/)
* [.Net Core, обмен с 1C по TCP/IP между различными устройствами](https://infostart.ru/public/551698/)

</details>

С одобрения [автора](https://infostart.ru/profile/82159/) разработка была выложена в этот репозиторий, но с некоторыми модификациями.

## Лицензия

MIT - делайте все, что посчитаете нужным.