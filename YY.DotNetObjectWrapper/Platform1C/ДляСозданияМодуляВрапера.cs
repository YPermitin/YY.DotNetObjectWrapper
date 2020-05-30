using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using YY.DotNetObjectWrapper.Templates;

namespace YY.DotNetObjectWrapper.Platform1C
{
    public static class ДляСозданияМодуляВрапера
    {
        #region Public Static Methods

        public static List<Tuple<string, Type>> ПолучитьПараметрыСТипами(EventInfo событие)
        {
            var rez = new List<Tuple<string, Type>>();
            var параметры = событие
                .EventHandlerType
                .GetMethod("Invoke")
                ?.GetParameters();

            if (параметры != null)
                foreach (var параметр in параметры)
                {
                    rez.Add(new Tuple<string, Type>(параметр.Name, параметр.ParameterType));
                }

            return rez;
        }
        public static void ПолучитьТекстСобытийСтипомПараметров(Type тип, List<Tuple<string, List<Tuple<string, Type>>>> res)
        {
            foreach (EventInfo e in тип.GetEvents())
            {
                var Метод = e.EventHandlerType.GetMethod("Invoke");

                if (Метод != null && Метод.ReturnType != typeof(void))
                {
                    continue;
                }
                res.Add(new Tuple<string, List<Tuple<string, Type>>>(e.Name, ПолучитьПараметрыСТипами(e)));
            }
        }
        public static string СоздатьОписаниеМодуляДля1С8(Type тип)
        {
            var стр = 
@"Перем врап,ОберткаСобытий;

Процедура СоздатьОбертку(объект)
ОберткаСобытий=врап.СоздатьОберткуДляСобытий(объект);

ДобавитьОбработчик ОберткаСобытий.ОшибкаСобытия,ОшибкаСобытия;

";

            StringBuilder ФункцияСозданияВрапера = new StringBuilder();
            ФункцияСозданияВрапера.AppendLine(стр);
            var СобытиеОшибки =
@"
//ИмяСобытия:String Имя События в котором произошло исключение
//Данные:object Параметры события
//ИсключениеСобытия:Exception Ошибка произошедшая при вызове события
Процедура ОшибкаСобытия(ИмяСобытия,Данные,ИсключениеСобытия)
	Сообщить(""Не обработано событие ""+ИмяСобытия);
     Сообщить("" Исключение ""+Врап.ВСтроку(ИсключениеСобытия));
 Сообщить("" Данные ""+Врап.ВСтроку(Данные));
            КонецПроцедуры  ";

            StringBuilder МетодыСобытий = new StringBuilder();
            //    StringBuilder ДляРеализацииСобытия = new StringBuilder();
            МетодыСобытий.AppendLine(СобытиеОшибки);
            var res = new List<Tuple<string, List<Tuple<string, Type>>>>();
            ПолучитьТекстСобытийСтипомПараметров(тип, res);

            foreach (var value in res)
            {

                ЗаполнитьФункцияСозданияВрапера(ФункцияСозданияВрапера, value);
                ЗаполнитьМетодыСобытий(МетодыСобытий, value);

            }

            ФункцияСозданияВрапера.AppendLine("КонецПроцедуры");
            ФункцияСозданияВрапера.AppendLine("");
            ФункцияСозданияВрапера.AppendLine(МетодыСобытий.ToString());

            return ФункцияСозданияВрапера.ToString();
        }
        public static string СоздатьОписаниеМодуляДля1C77(Type тип)
        {
            var стр = 
@" 
Перем врап,ОберткаСобытий;

           Функция СоздатьОбертку(ОбертываемыйОбъект)
            ПодключитьВнешнююКомпоненту(""AddIn.GlobalContext1C"");
            объект = СоздатьОбъект(""AddIn.GlobalContext1C"");
            ГлобальныйКонтекст = объект.ГлобальныйКонтекст;

            ОберткаСобытий = врап.СоздатьОберткуДляСобытий77(ОбертываемыйОбъект,ГлобальныйКонтекст);
           КонецФункции // СоздатьОбертку
";
            StringBuilder ФункцияСозданияВрапера = new StringBuilder();
            ФункцияСозданияВрапера.AppendLine(стр);
            var СобытиеОшибки = 
@"
// Свойства ОберткаСобытий.ПоследняяОшибка
//Событие:String Имя События в котором произошло исключение
//Данные:object Параметры события
//ИсключениеСобытия:Exception Ошибка произошедшая при вызове события
Функция ОшибкаСобытия()
	ПоследняяОшибка=ОберткаСобытий.ПоследняяОшибка;
	Сообщить(""Не обработано событие ""+ПоследняяОшибка.Событие);
    Сообщить(Врап.ВСтроку(Шаблон(""[ОберткаСобытий."" + ПоследняяОшибка.Событие + ""]"")));
    Сообщить(""Ошибка"");
    Сообщить(врап.ВСтроку(ПоследняяОшибка.Исключение))
 КонецФункции  ";

            StringBuilder МетодыСобытий = new StringBuilder();
            МетодыСобытий.AppendLine(СобытиеОшибки);
            var res = new List<Tuple<string, List<Tuple<string, Type>>>>();
            ПолучитьТекстСобытийСтипомПараметров(тип, res);

            foreach (var value in res)
            {
                ЗаполнитьМетодыСобытий77(МетодыСобытий, value);
            }
            string ТипСтрРеальногоОбъекта = тип.FullName;
            if (ТипСтрРеальногоОбъекта != null)
            {
                var ИмяКласса = ТипСтрРеальногоОбъекта.Replace(".", "_").Replace("+", "_");
                string Заключение = 
                    @"

Процедура ПриОткрытии()
	врап=СоздатьОбъект(""NetObjectToIDispatch45"");

КонецПроцедуры // ПриОткрытии
               //======================================================================
Процедура ОбработкаВнешнегоСобытия(Источник, ИмяСобытия, Данные)
             Если Источник = ""{0}"" Тогда

                  Шаблон(""["" + ИмяСобытия + ""()]"");
            КонецЕсли;
            КонецПроцедуры // ОбработкаВнешнегоСобытия  ";
                ФункцияСозданияВрапера.AppendLine("");
                ФункцияСозданияВрапера.AppendLine(МетодыСобытий.ToString());
                ФункцияСозданияВрапера.AppendLine(string.Format(Заключение, ИмяКласса));
            }

            return ФункцияСозданияВрапера.ToString();
        }
        public static string СоздатьОписания(Type тип)
        {

            StringBuilder ДляИнтерфейса = new StringBuilder();
            StringBuilder ДляОписанияСобытия = new StringBuilder();
            StringBuilder ДляРеализацииСобытия = new StringBuilder();

            var res = new List<Tuple<string, List<string>>>();
            ПолучитьТекстСобытий(тип, res);
            int i = 1;
            foreach (var value in res)
            {
                i++;
                ЗаполнитьДляИнтерфейса(ДляИнтерфейса, value, i);
                ЗаполнитьОписанияСобытий(ДляОписанияСобытия, value);
                ЗаполнитьРеализацииСобытий(ДляРеализацииСобытия, value);
            }


            var СобытияИнтерфейса = ДляИнтерфейса.ToString();
            var ОписанияСобытия = ДляОписанияСобытия.ToString();
            var РеализацииСобытий = ДляРеализацииСобытия.ToString();
            string ТипРеальногоОбъекта = тип.FullName;
            if (ТипРеальногоОбъекта != null)
            {
                var ИмяКласса = ТипРеальногоОбъекта.Replace(".", "_").Replace("+", "_");
                var ГуидИнтерфейса = Guid.NewGuid().ToString();
                var ГуидКласса = Guid.NewGuid().ToString();

                var СтрокаМодуля = string.Format(ClassTemplates.Template, ГуидИнтерфейса, ИмяКласса, СобытияИнтерфейса, ГуидКласса, ТипРеальногоОбъекта, ОписанияСобытия, РеализацииСобытий);
                return СтрокаМодуля;
            }

            return null;
        }
        public static void ПолучитьТекстСобытий(Type тип, List<Tuple<string, List<string>>> res)
        {
            bool добавлять;
            foreach (EventInfo e in тип.GetEvents())            
            {
                var параметры = ПолучитьПараметры(e, out добавлять);
                if (добавлять) res.Add(new Tuple<string, List<string>>(e.Name, параметры));
            }
        }
        public static List<string> ПолучитьПараметры(EventInfo событие, out bool добавлять)
        {
            добавлять = true;
            var rez = new List<string>();
            var Метод = событие.EventHandlerType.GetMethod("Invoke");

            if (Метод != null && Метод.ReturnType != typeof(void))
            {
                добавлять = false;
                return rez;
            }

            if (Метод != null)
            {
                var параметры = Метод.GetParameters();
                foreach (var параметр in параметры)
                {
                    rez.Add(параметр.Name);
                }
            }

            return rez;
        }
        public static string СоздатьОписания77(Type тип)
        {


            StringBuilder ДляОписанияСобытия = new StringBuilder();
            StringBuilder ДляРеализацииСобытия = new StringBuilder();

            var res = new List<Tuple<string, List<string>>>();
            ПолучитьТекстСобытий(тип, res);
            foreach (var value in res)
            {
                ЗаполнитьОписанияСобытий77(ДляОписанияСобытия, value);
                ЗаполнитьРеализацииСобытий77(ДляРеализацииСобытия, value);
            }

            var ОписанияСобытия = ДляОписанияСобытия.ToString();
            var РеализацииСобытий = ДляРеализацииСобытия.ToString();
            string ТипРеальногоОбъекта = тип.FullName;
            if (ТипРеальногоОбъекта != null)
            {
                var ИмяКласса = ТипРеальногоОбъекта.Replace(".", "_").Replace("+", "_");
                var СтрокаМодуля = string.Format(ClassTemplates.Template77, ИмяКласса, ТипРеальногоОбъекта, ОписанияСобытия, РеализацииСобытий);

                return СтрокаМодуля;
            }

            return null;
        }

        #endregion

        #region Private Static Methods

        private static void ПолучитьОписаниеПараметров(StringBuilder функцияСозданияВрапера, List<Tuple<string, Type>> value)
        {
            if (value.Count == 0)
                return;

            string str;
            if (value.Count == 1)
            {
                str = @"//  параметр Данные:{0}";
                функцияСозданияВрапера.AppendLine(string.Format(str, value[0].Item2.FullName));
            }
            else
            {
                str = @"//  параметр Данные:Аннимный Тип
                       // Свойства параметра";
                функцияСозданияВрапера.AppendLine(str);
                foreach (var свойство in value)
                {
                    str = "// {0}:{1}";
                    функцияСозданияВрапера.AppendLine(string.Format(str, свойство.Item1, свойство.Item2.FullName));
                }
            }
        }
        private static void ЗаполнитьМетодыСобытий(StringBuilder методыСобытий, Tuple<string, List<Tuple<string, Type>>> value)
        {
            string событие = value.Item1;

            string Данные = value.Item2.Count > 0 ? "Данные" : "";
            string ТелоПроцедуры = value.Item2.Count > 0 ? @"Сообщить(""" + событие + @" ""+Врап.ВСтроку(Данные));" : @"Сообщить(""" + событие + @""")";
            string стр = 
@"
Процедура {0}({1})
    {2}
КонецПроцедуры";

            ПолучитьОписаниеПараметров(методыСобытий, value.Item2);
            методыСобытий.AppendLine(string.Format(стр, событие, Данные, ТелоПроцедуры));
        }
        private static void ЗаполнитьФункцияСозданияВрапера(StringBuilder функцияСозданияВрапера, Tuple<string, List<Tuple<string, Type>>> value)
        {
            string событие = value.Item1;
            функцияСозданияВрапера.AppendLine(string.Format("ДобавитьОбработчик ОберткаСобытий.{0}, {0};", событие));
        }
        private static void ПолучитьОписаниеПараметров77(string имясобытия, StringBuilder функцияСозданияВрапера, List<Tuple<string, Type>> value)
        {
            if (value.Count == 0)
                return;

            string str;
            if (value.Count == 1)
            {
                str = @"// ОберткаСобытий.{0}:{1}";
                функцияСозданияВрапера.AppendLine(string.Format(str, имясобытия, value[0].Item2.FullName));
            }
            else
            {
                str = @"//  Свойства ОберткаСобытий.{0}";
                функцияСозданияВрапера.AppendLine(string.Format(str, имясобытия));
                foreach (var свойство in value)
                {
                    str = "// {0}:{1}";
                    функцияСозданияВрапера.AppendLine(string.Format(str, свойство.Item1, свойство.Item2.FullName));
                }
            }
        }
        private static void ЗаполнитьМетодыСобытий77(StringBuilder методыСобытий, Tuple<string, List<Tuple<string, Type>>> value)
        {
            string событие = value.Item1;

            string ТелоПроцедуры = string.Format(value.Item2.Count > 0 ? @"Сообщить(""{0} ""+Врап.ВСтроку(ОберткаСобытий.{0}));" : @"Сообщить(""{0}"")", событие);
            string стр = @"
            Функция {0}()
               {1}
            КонецФункции
";
            ПолучитьОписаниеПараметров77(событие, методыСобытий, value.Item2);
            методыСобытий.AppendLine(string.Format(стр, событие, ТелоПроцедуры));
        }
        private static void ЗаполнитьОписанияСобытий77(StringBuilder дляОписанияСобытия, Tuple<string, List<string>> value)
        {
            string ИмяСобытия = value.Item1;
            var Параметры = value.Item2;
            if (Параметры.Count == 0)
                return;

            string шаблон = @"public object {0};";
            дляОписанияСобытия.AppendLine(string.Format(шаблон, ИмяСобытия));

        }
        private static void ЗаполнитьРеализацииСобытий77(StringBuilder дляРеализацииСобытия, Tuple<string, List<string>> value)
        {
            string ИмяСобытия = value.Item1;
            var Параметры = value.Item2;

            if (Параметры.Count == 0)
            {
                var str = 
@"РеальныйОбъект.{0} += () =>
{{
    ОтослатьСобытие(""{0}"");
}};";

                дляРеализацииСобытия.AppendLine(string.Format(str, ИмяСобытия));
                return;
            }
            else if (Параметры.Count == 1)
            {

                var str = 
@"РеальныйОбъект.{0} += ({1}) =>
{{
    {0} = {1};
    ОтослатьСобытие(""{0}"");
}};";

                String ИмяПараметра = Параметры[0];
                дляРеализацииСобытия.AppendLine(string.Format(str, ИмяСобытия, ИмяПараметра));
                return;
            }

            StringBuilder ПараметрыСобытия = new StringBuilder();
            StringBuilder СвойстваКласса = new StringBuilder();
            foreach (var Параметр in Параметры)
            {
                ПараметрыСобытия.Append(Параметр + ",");
                СвойстваКласса.Append(Параметр + "=" + Параметр + ",");

            }

            string шаблон = 
@"РеальныйОбъект.{0} += ({2}) =>
{{
    {0} =  new {{ {1}}};
    ОтослатьСобытие(""{0}"");
}};
";

            string strClass = СвойстваКласса.ToString(0, СвойстваКласса.Length - 1);
            string strParam = ПараметрыСобытия.ToString(0, ПараметрыСобытия.Length - 1);
            дляРеализацииСобытия.AppendLine(string.Format(шаблон, ИмяСобытия, strClass, strParam));
        }
        private static void ЗаполнитьРеализацииСобытий(StringBuilder дляРеализацииСобытия, Tuple<string, List<string>> value)
        {
            string ИмяСобытия = value.Item1;
            var Параметры = value.Item2;

            if (Параметры.Count == 0)
            {
                var str = 
@"РеальныйОбъект.{0} += () =>
{{
    ОтослатьСобытие({0}, ""{0}"");
}};";

                дляРеализацииСобытия.AppendLine(string.Format(str, ИмяСобытия));
                return;
            }
            else if (Параметры.Count == 1)
            {

                var str = 
@"РеальныйОбъект.{0} += ({1}) =>
{{
    ОтослатьСобытиеСПараметром({0}, {1}, ""{0}"");
}};";

                String ИмяПараметра = Параметры[0];
                дляРеализацииСобытия.AppendLine(string.Format(str, ИмяСобытия, ИмяПараметра));
                return;
            }

            StringBuilder ПараметрыСобытия = new StringBuilder();
            StringBuilder СвойстваКласса = new StringBuilder();
            foreach (var Параметр in Параметры)
            {
                ПараметрыСобытия.Append(Параметр + ",");
                СвойстваКласса.Append(Параметр + "=" + Параметр + ",");

            }

            string шаблон = 
@"РеальныйОбъект.{0} += ({2}) =>
{{
    ОтослатьСобытиеСПараметром({0}, new {{ {1}}}, ""{0}"");
}};
";

            string strClass = СвойстваКласса.ToString(0, СвойстваКласса.Length - 1);
            string strParam = ПараметрыСобытия.ToString(0, ПараметрыСобытия.Length - 1);
            дляРеализацииСобытия.AppendLine(string.Format(шаблон, ИмяСобытия, strClass, strParam));
        }
        private static void ЗаполнитьДляИнтерфейса(StringBuilder дляИнтерфейса, Tuple<string, List<string>> value, int i)
        {
            var ДиспИдШаблон = "0x00000001";
            var ДлинаШаблона = ДиспИдШаблон.Length;

            var СтрИд = i.ToString();
            var ДиспИд = ДиспИдШаблон.Substring(0, ДлинаШаблона - СтрИд.Length) + СтрИд;
            дляИнтерфейса.AppendLine(string.Format(@"[DispId({0})]", ДиспИд));

            string ИмяСобытия = value.Item1;
            var Параметры = value.Item2;
            string strParam = "object value";
            if (Параметры.Count == 0)
                strParam = "";

            дляИнтерфейса.AppendLine(string.Format(@"void {0}({1});", ИмяСобытия, strParam));
        }
        private static void ЗаполнитьОписанияСобытий(StringBuilder дляОписанияСобытия, Tuple<string, List<string>> value)
        {
            string ИмяСобытия = value.Item1;
            var Параметры = value.Item2;
            string strParam = "СобытиеСПараметром_Delgate";
            if (Параметры.Count == 0)
                strParam = "Событие_Delgate";

            string шаблон = @"public event {0} {1};";
            дляОписанияСобытия.AppendLine(string.Format(шаблон, strParam, ИмяСобытия));
        }

        #endregion
    }
}
