# DotNet Object Wrapper

Оболочка объектов .NET для использования в других экосистемах приложений.

## Предыстория

Изначально проект назывался **["NetObjetToIDispatch"](https://infostart.ru/public/238584/)** и создавался силами **[Сергея Смирнова](https://infostart.ru/profile/82159/)**.

Данная разработка создавалась для использования сборок .NET в решениях на платформе 1С:Предприятие через преобразование объектов и классов в COM-объекты. Последние можно использовать в экосистеме 1С. Достигается это путем создания класса, реализующего методы интерфейса **[IReflect](https://docs.microsoft.com/ru-ru/dotnet/api/system.reflection.ireflect?view=netcore-3.1)**.