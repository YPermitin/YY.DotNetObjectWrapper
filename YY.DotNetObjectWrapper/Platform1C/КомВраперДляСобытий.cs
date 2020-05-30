using Microsoft.CSharp;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using YY.DotNetObjectWrapper.Service;
using YY.DotNetObjectWrapper.Templates;

namespace YY.DotNetObjectWrapper.Platform1C
{
    [ComVisible(false)]
    public class КомВраперДляСобытий<T>
    {
        #region Public Static Members

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
                var ГуидИнтерфейса = Guid.NewGuid().ToString(); //0 Гуид Интерфейса 
                //1 ИмяКласса врапера
                //2 Описание Методов в Интерфейсе
                var ГуидКласса = Guid.NewGuid().ToString();       //3 Гуид Класса
                //4 Тип Реального объекта
                //5 Реализация событий
                //

                var СтрокаМодуля = string.Format(ClassTemplates.Template, ГуидИнтерфейса, ИмяКласса, СобытияИнтерфейса, ГуидКласса, ТипРеальногоОбъекта, ОписанияСобытия, РеализацииСобытий);
                return СтрокаМодуля;
            }
            else
            {
                return null;
            }
        }

        #endregion

        #region Private Static Members

        private static MethodInfo _cоздательОбертки;
        private static MethodInfo _cоздательОбертки77;

        #endregion

        #region Public Static methods

        public static void ПолучитьТекстСобытий(Type тип, List<Tuple<string, List<string>>> res)
        {
            foreach (EventInfo e in тип.GetEvents())
            {
                var параметры = ПолучитьПараметры(e, out var добавлять);

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

        #endregion

        #region Private Static Methods

        private static void ДобавитьссылкинаDllвКаталоге(Type тип, System.Collections.Specialized.StringCollection referencedAssemblies)
        {
            string ИмяФайлаСборки = тип.Assembly.Location;
            string Каталог = Path.GetDirectoryName(ИмяФайлаСборки);
            string[] files = Directory.GetFiles(Каталог ?? string.Empty, "*.dll");

            foreach (string f in files)
            {
                if (!String.Equals(f, ИмяФайлаСборки, StringComparison.CurrentCultureIgnoreCase))
                    if (referencedAssemblies.IndexOf(f) == -1)
                        referencedAssemblies.Add(f);
            }
        }
        private static CompilerResults СкомпилироватьОбертку(string строкаКласса, string имяКласса)
        {
            bool ЭтоСборкаГак = typeof(T).Assembly.GlobalAssemblyCache;
            string Путь = Path.GetDirectoryName(typeof(T).Assembly.Location);

            string OutputAssembly = Path.Combine(Путь ?? string.Empty, имяКласса) + ".dll";
            var compiler = new CSharpCodeProvider();
            var parameters = new CompilerParameters();
            parameters.ReferencedAssemblies.Add("System.dll");
            parameters.ReferencedAssemblies.Add("System.Core.dll");
            parameters.ReferencedAssemblies.Add("Microsoft.CSharp.dll");
            parameters.ReferencedAssemblies.Add("System.Windows.Forms.dll");
            parameters.ReferencedAssemblies.Add(typeof(AutoWrap).Assembly.Location);

            if (!ЭтоСборкаГак)
            {
                parameters.ReferencedAssemblies.Add(typeof(T).Assembly.Location);
                ДобавитьссылкинаDllвКаталоге(typeof(T), parameters.ReferencedAssemblies);
            }
            else
            {
                string ИмяСборки = typeof(T).Assembly.ManifestModule.Name;
                if (parameters.ReferencedAssemblies.IndexOf(ИмяСборки) == -1)
                    parameters.ReferencedAssemblies.Add(ИмяСборки);
            }

            if (ЭтоСборкаГак)
                parameters.GenerateInMemory = true;
            else
            {
                parameters.GenerateInMemory = false;
                parameters.OutputAssembly = OutputAssembly;
            }

            parameters.GenerateExecutable = false;
            parameters.IncludeDebugInformation = true;

            var res = compiler.CompileAssemblyFromSource(parameters, строкаКласса);
            return res;
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

        #endregion

        #region Public Properties

        public static MethodInfo СоздательОбертки
        {
            get
            {
                if (_cоздательОбертки == null)
                {
                    Type типРеальногоОбъекта = typeof(T);
                    string ТипСтрРеальногоОбъекта = типРеальногоОбъекта.FullName;
                    if (ТипСтрРеальногоОбъекта != null)
                    {
                        var ИмяКласса = "ВрапперДля" + ТипСтрРеальногоОбъекта.Replace(".", "_").Replace("+", "_");
                        string строкаКласса = ДляСозданияМодуляВрапера.СоздатьОписания(типРеальногоОбъекта);
                        var res = СкомпилироватьОбертку(строкаКласса, ИмяКласса);

                        Type тип = res.CompiledAssembly.GetType(ИмяКласса);
                        MethodInfo mi = тип.GetMethod("СоздатьОбъект", new[] { typeof(object), типРеальногоОбъекта });

                        _cоздательОбертки = mi;
                    }
                }

                return _cоздательОбертки;
            }
        }
        public static MethodInfo СоздательОбертки77
        {
            get
            {
                if (_cоздательОбертки77 == null)
                {
                    Type типРеальногоОбъекта = typeof(T);
                    string ТипСтрРеальногоОбъекта = типРеальногоОбъекта.FullName;
                    if (ТипСтрРеальногоОбъекта != null)
                    {
                        var ИмяКласса = "ВрапперДля" + ТипСтрРеальногоОбъекта.Replace(".", "_").Replace("+", "_") + "77";
                        string строкаКласса = ДляСозданияМодуляВрапера.СоздатьОписания77(типРеальногоОбъекта);
                        var res = СкомпилироватьОбертку(строкаКласса, ИмяКласса);

                        Type тип = res.CompiledAssembly.GetType(ИмяКласса);
                        MethodInfo mi = тип.GetMethod("СоздатьОбъект", new[] { typeof(object), типРеальногоОбъекта });

                        _cоздательОбертки77 = mi;
                    }
                }

                return _cоздательОбертки77;
            }
        }

        #endregion
    }
}
