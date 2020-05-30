using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Configuration;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ServiceModel.Configuration;
using System.Windows.Forms;
using YY.DotNetObjectWrapper.Platform1C;
using YY.DotNetObjectWrapper.Service;

namespace YY.DotNetObjectWrapper
{
    [ComVisible(true)]
    [ProgId("YY.DotNetObjectWrapper")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("DFDADA57-B22C-4276-928A-8B91C9891FF1")]
    public class DotNetObjectWrapper
    {
        #region Static Constructor

        static DotNetObjectWrapper()
        {
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.AssemblyResolve += MyResolveEventHandler;
        }

        #endregion

        #region Private Static Methods

        private static Type НайтиТип(string type)
        {
            Type result = Type.GetType(type, false);

            if (result != null)
            {

                return result;
            }

            foreach (Assembly tmpAasembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                result = tmpAasembly.GetType(type);
                if (result != null)
                {
                    return result;
                }
            }

            return null;
        }
        private static Assembly MyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            var FullName = args.Name;
            foreach (Assembly tmpAasembly in AppDomain.CurrentDomain.GetAssemblies())
            {

                if (tmpAasembly.FullName == FullName)
                {
                    return tmpAasembly;
                }
            }

            return typeof(DotNetObjectWrapper).Assembly;
        }

        #endregion

        #region Private Properies

        public object Null { get { return null; } }

        #endregion

        #region Public Properties

        public object Activator
        {
            get { return new AutoWrap(typeof(Activator)); }
        }
        public bool ВыводитьСообщениеОбОшибке
        {
            get { return AutoWrap.ВыводитьСообщениеОбОшибке; }
            set { AutoWrap.ВыводитьСообщениеОбОшибке = value; }
        }
        public object ПоследняяОшибка
        {
            get
            {
                if (AutoWrap.ПоследняяОшибка == null) return null;

                return new AutoWrap(AutoWrap.ПоследняяОшибка);
            }
        }

        #endregion

        #region Public Methods

        public void ЗакрытьРесурс(Object oбъект)
        {
            object объект = AutoWrap.ПолучитьРеальныйОбъект(oбъект);

            if (объект is IDisposable d) d.Dispose();
        }
        public object СоздатьДелегат(object типДелегата, object методИнфо, Object oбъект, params object[] argOrig)
        {
            object объект = AutoWrap.ПолучитьРеальныйОбъект(oбъект);
            MethodInfo объектМетодИнфо = (MethodInfo)AutoWrap.ПолучитьРеальныйОбъект(методИнфо);
            Type genType = (Type)AutoWrap.ПолучитьРеальныйОбъект(типДелегата);
            var типы = AutoWrap.ПолучитьМассивРеальныхОбъектов(argOrig).Cast<Type>().ToArray();
            Type constructed = genType.MakeGenericType(типы);

            var res = Delegate.CreateDelegate(constructed, объект, объектМетодИнфо);
            return new AutoWrap(res);
        }
        public object ПолучитьТипизированныйПеречислитель(object счетчик, object тип)
        {


            var cur = (IEnumerable)AutoWrap.ПолучитьРеальныйОбъект(счетчик);

            Type type = (Type)AutoWrap.ПолучитьРеальныйОбъект(тип);
            var res = new ТипизированныйЭнумератор(cur.GetEnumerator(), type);
            return new MyEnumerableClass(res);
        }
        public object ПолучитьПеречислитель(object счетчик)
        {

            var cur = (IEnumerable)AutoWrap.ПолучитьРеальныйОбъект(счетчик);

            return new AutoWrap(cur.GetEnumerator(), typeof(IEnumerator));
        }
        public bool Следующий(object перечислитель, out object current)
        {
            var enumerator = (IEnumerator)AutoWrap.ПолучитьРеальныйОбъект(перечислитель);

            if (enumerator.MoveNext())
            {
                current = AutoWrap.ОбернутьОбъект(enumerator.Current);
                return true;
            }

            current = null;
            return false;
        }
        public object[] МассивИзЭнумератора(object счетчик, object тип)
        {
            var cur = (IEnumerable)AutoWrap.ПолучитьРеальныйОбъект(счетчик);

            var list = new List<object>();
            Type type = (Type)AutoWrap.ПолучитьРеальныйОбъект(тип);


            ДанныеДляТипа данныеДляТипа = ДанныеДляТипа.ПолучитьДанныеДляТипа(type);




            foreach (var str in cur)
            {


                if ((str == null) || !type.IsAssignableFrom(str.GetType()))
                {
                    list.Add(null);
                    continue;
                }


                var res = new AutoWrap(str, type);
                ДанныеДляТипа.ПрописатьПоля(res, данныеДляТипа);
                list.Add(res);
            }
            return list.ToArray();
        }
        public object[] МассивИзЭнумератора2(object счетчик)
        {
            var cur = (IEnumerable)AutoWrap.ПолучитьРеальныйОбъект(счетчик);

            var list = new List<object>();

            foreach (var str in cur)
            {

                list.Add(AutoWrap.ОбернутьОбъект(str));
            }
            return list.ToArray();
        }
        // Получаем массив объектов используя эффект получения массива параметров при использовании 
        // ключевого слова params
        // Пример использования 
        //МассивПараметров=ъ(Врап.Массив(doc.ПолучитьСсылку(), "a[title='The Big Bang Theory']"));
        public object Массив(params object[] argOrig)
        {
            argOrig = AutoWrap.ПолучитьМассивРеальныхОбъектов(argOrig);
            return AutoWrap.ОбернутьОбъект(argOrig);
        }
        // Получаем масив элементов определенного типа
        // Тип выводится по первому элементу с индексом 0
        // Пример использования
        // ТипыПараметров=ъ(Врап.ТипМассив(IParentNode.ПолучитьСсылку(),String.ПолучитьСсылку()));
        public object ТипМассив(params object[] argOrig)
        {
            if (!(argOrig != null && argOrig.Length > 0))
                return new object[0];

            argOrig = AutoWrap.ПолучитьМассивРеальныхОбъектов(argOrig);
            var Тип = argOrig[0].GetType();
            var ti = Тип.GetTypeInfo();
            var TypeRes = typeof(List<>).MakeGenericType(Тип);

            var res = System.Activator.CreateInstance(TypeRes);
            var mi = TypeRes.GetTypeInfo().GetMethod("Add");
            var mi2 = TypeRes.GetTypeInfo().GetMethod("ToArray");
            foreach (var str in argOrig)
            {
                if (mi != null && str != null && ti.IsInstanceOfType(str))
                    mi.Invoke(res, new[] { str });

            }

            if(mi2 != null)
                return AutoWrap.ОбернутьОбъект(mi2.Invoke(res, new object[0]));
            else
                return null;
        }
        public object ПолучитьАсинхронныйВыполнитель()
        {
            return new АсинхронныйВыполнитель();
        }
        public void УстЭтоСемерка()
        {
            AutoWrap.ЭтоСемерка = true;
        }
        public object ОбернутьЛюбойОбъект(object объект)
        {
            if (объект == null) return null;

            if (объект is AutoWrap)
                return объект;

            return new AutoWrap(объект);
        }
        public object ОбернутьОбъект(object объект)
        {
            if (объект is AutoWrap)
                return объект;

            return AutoWrap.ОбернутьОбъект(объект);
        }
        public object ПолучитьРеальныйОбъект(object valueOrig)
        {
            return AutoWrap.ПолучитьРеальныйОбъект(valueOrig);
        }
        public string ВСтроку(object valueOrig)
        {
            if (valueOrig == null)
                return "неопределено";

            var res = AutoWrap.ПолучитьРеальныйОбъект(valueOrig);
            return res.ToString();
        }
        public void ВывестиПоследнююОшибку()
        {
            var e = AutoWrap.ПоследняяОшибка;
            string Ошибка = "Нет ошибки";

            if (e != null)
            {
                Ошибка = e.Message + " " + e.Source;

                if (e.InnerException != null)
                    Ошибка = Ошибка + "\r\n" + e.InnerException;
            }


            MessageBox.Show(Ошибка);
            if (e != null)
                MessageBox.Show(e.StackTrace);
        }
        public object ПолучитьОбъектДляСобытий(object объект, string имяСобытия)
        {
            var res = new ClassForEvent1C(((AutoWrap)объект).O, имяСобытия);
            return res;
        }
        public object ПолучитьОбъектДляСобытийСПараметром(object объект, string имяСобытия)
        {
            var res = new ClassForEvent1C(((AutoWrap)объект).O, имяСобытия, true);
            return res;
        }
        public void ОчиститьСобытияОбъекта(Object obj)
        {
            if (obj == null)
                return;

            ОчисткаСсылокНаСобытия.Очистить(AutoWrap.ПолучитьРеальныйОбъект(obj));
        }
        public string УстановитьDefaultProxy()
        {
            Configuration roamingConfig = ConfigurationManager
                .OpenExeConfiguration(ConfigurationUserLevel.None);

            ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap
            {
                ExeConfigFilename = roamingConfig.FilePath
            };

            // Get the mapped configuration file.
            Configuration config =
              ConfigurationManager.OpenMappedExeConfiguration(
                configFileMap, ConfigurationUserLevel.None);

            string sectionName = "system.net/defaultProxy";

            if (!(config.GetSection(sectionName) is DefaultProxySection ProxySection))
            {
                ProxySection = new DefaultProxySection();
                config.Sections.Add(sectionName, ProxySection);
            }
            ProxySection.UseDefaultCredentials = true;
            ProxySection.Enabled = true;

            config.Save(ConfigurationSaveMode.Full, true);
            ConfigurationManager.RefreshSection(sectionName);

            return config.FilePath;
        }
        public object СоздатьКлиентаWcfConfigFile(string имяФайла, object channel, string endpointConfigurationName, object endpointAddress = null, string userName = null, string password = null)
        {
            ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap {ExeConfigFilename = имяФайла};
            Configuration newConfiguration = ConfigurationManager.OpenMappedExeConfiguration(
                fileMap,
                ConfigurationUserLevel.None);

            Type ТипКанала = ТипДляСоздатьОбъект(channel);
            Type type = typeof(ConfigurationChannelFactory<>);
            Type constructed = type.MakeGenericType(ТипКанала);

            dynamic factory1 = System.Activator.CreateInstance(constructed,
                endpointConfigurationName,
                newConfiguration,
                AutoWrap.ПолучитьРеальныйОбъект(endpointAddress)
                );

            if (!String.IsNullOrWhiteSpace(userName))
            {

                factory1.Credentials.UserName.UserName = userName;
                factory1.Credentials.UserName.Password = password;


            }
            
            return AutoWrap.ОбернутьОбъект(factory1.CreateChannel());
        }
        public void ЗаменитьConfigFile(string имяФайла)
        {
            if (string.IsNullOrEmpty(имяФайла))
                имяФайла = typeof(DotNetObjectWrapper).Assembly.Location + ".config";

            AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", имяФайла);

            typeof(ConfigurationManager)
                .GetField("s_initState", BindingFlags.NonPublic |
                                         BindingFlags.Static)
                ?.SetValue(null, 0);

            typeof(ConfigurationManager)
                .GetField("s_configSystem", BindingFlags.NonPublic |
                                            BindingFlags.Static)
                ?.SetValue(null, null);
        }
        public object СоздатьОберткуДляСобытий(object объект)
        {
            object РеальныйОбъект = AutoWrap.ПолучитьРеальныйОбъект(объект);
            var функция = ПолучитьMethodInfoОберткиСобытий("СоздательОбертки", РеальныйОбъект);
            object обертка = функция.Invoke(null, new[] { this, РеальныйОбъект });

            return обертка;
        }
        public object СоздатьОберткуДляСобытий77(object объект, object глобальныйКонтекст)
        {
            object РеальныйОбъект = AutoWrap.ПолучитьРеальныйОбъект(объект);
            var функция = ПолучитьMethodInfoОберткиСобытий("СоздательОбертки77", РеальныйОбъект);
            object обертка = функция.Invoke(null, new[] { глобальныйКонтекст, РеальныйОбъект });

            return AutoWrap.ОбернутьОбъект(обертка);
        }
        public object CreateObject(object type)
        {
            var res = ТипДляСоздатьОбъект(type);
            return AutoWrap.ОбернутьОбъект(System.Activator.CreateInstance(res));
        }
        public object СоздатьОбъект(object тип, params object[] argOrig)
        {
            var res = ТипДляСоздатьОбъект(тип);

            object[] args = AutoWrap.ПолучитьМассивРеальныхОбъектов(argOrig);
            return AutoWrap.ОбернутьОбъект(System.Activator.CreateInstance(res, args));
        }
        public object CreateObjectWhithParam(object тип, Object[] параметры)
        {
            return СоздатьОбъект(тип, параметры);
        }
        public object CreateArray(object type, int length)
        {
            return AutoWrap.ОбернутьОбъект(Array.CreateInstance(ТипДляСоздатьОбъект(type), length));
        }
        public object СоздатьМассив(object type, int length)
        {
            return CreateArray(type, length);
        }
        public object GetValueFromArray(object массивOrig, Object индекс)
        {
            object Массив = массивOrig;

            if (Массив is AutoWrap)
                Массив = ((AutoWrap)Массив).O;

            int ind = Convert.ToInt32(индекс);

            Array Масс = (Array)Массив;

            return AutoWrap.ОбернутьОбъект(Масс.GetValue(ind));
        }
        public object ПолучитьТип(string type)
        {

            var res = НайтиТип(type);

            if (res == null)
            {
                string ошибка = " неверный тип " + type;
                MessageBox.Show(ошибка);
                throw new COMException(ошибка);
            }

            return new AutoWrap(res);

        }
        public object ПолучитьТипИзСборки(string type, string путь)
        {
            var result = НайтиТип(type);
            if (result != null)
                return new AutoWrap(result);

            string ошибка;
            string path;
            Assembly assembly;
            if (File.Exists(путь))
                assembly = Assembly.LoadFrom(путь);
            else
            {
                path = Path.Combine(Path.GetDirectoryName(typeof(string).Assembly.Location) ?? string.Empty, путь);
                if (File.Exists(path))
                {
                    var ИмяСборки = AssemblyName.GetAssemblyName(path).FullName;
                    assembly = Assembly.Load(ИмяСборки);
                }
                else
                {
                    ошибка = " Не найден файл " + путь + " или  " + path;
                    MessageBox.Show(ошибка);
                    throw new COMException(ошибка);
                }
            }

            result = assembly.GetType(type);
            if (result != null)
            {

                return new AutoWrap(result);
            }

            ошибка = " неверный тип " + type + " в сборке " + assembly.Location;

            MessageBox.Show(ошибка);
            throw new COMException(ошибка);
        }
        public Object ChangeType(string type, object valueOrig)
        {
            object value = AutoWrap.ПолучитьРеальныйОбъект(valueOrig);

            Type result = НайтиТип(type);
            if (result == null)
                throw new COMException("Не найден тип " + type);

            return new AutoWrap(Convert.ChangeType(value, result, CultureInfo.InvariantCulture));
        }
        public Object ToDecimal(object valueOrig)
        {
            object value = AutoWrap.ПолучитьРеальныйОбъект(valueOrig);

            return new AutoWrap(Convert.ChangeType(value, typeof(Decimal), CultureInfo.InvariantCulture));
        }
        public Object ToInt(object valueOrig)
        {
            object value = AutoWrap.ПолучитьРеальныйОбъект(valueOrig);

            return new AutoWrap(Convert.ChangeType(value, typeof(Int32), CultureInfo.InvariantCulture));
        }
        public object ЗагрузитьСборку(string путьКСборке)
        {
            var res = Assembly.LoadFrom(путьКСборке);
            return AutoWrap.ОбернутьОбъект(res);
        }
        public void GetInfoFromObject(Object obj)
        {
            AutoWrap res = (AutoWrap)obj;

            MessageBox.Show(res.T.AssemblyQualifiedName);
            MessageBox.Show(res.O.ToString());

            foreach (MethodInfo str in res.Методы)
            {
                MessageBox.Show(str.Name);
            }
        }
        public object ПолучитьИнтерфейс(object obj, string interfaseName)
        {
            object РеальныйОбъект = AutoWrap.ПолучитьРеальныйОбъект(obj);

            Type type = РеальныйОбъект.GetType().GetInterface(interfaseName, true);
            if (type == null) return null;

            return new AutoWrap(РеальныйОбъект, type);
        }
        public object ПолучитьМассивИзЛистаСтрок(object лист)
        {
            return new AutoWrap(((List<String>)(((AutoWrap)лист).O)).ToArray());
        }
        public object ТипКакОбъект(object тип)
        {
            if (тип is AutoWrap)
                тип = ((AutoWrap)тип).T;
            Type T = ((Type)тип);

            return new AutoWrap(T, T);
        }
        public object ПолучитьМетоды(object тип)
        {
            object РеальныйОбъект = AutoWrap.ПолучитьРеальныйОбъект(тип);

            return new AutoWrap(((Type)РеальныйОбъект).GetMethods());
        }
        public object ВыполнитьДелегат(object делегат, params object[] argsOrig)
        {
            Delegate obj = (Delegate)ПолучитьРеальныйОбъект(делегат);
            object[] args = AutoWrap.ПолучитьМассивРеальныхОбъектов(argsOrig);

            return AutoWrap.ОбернутьОбъект(obj.DynamicInvoke(args));
        }
        public object ВыполнитьМетод(object objOrig, string имяМетода, params object[] argsOrig)
        {
            object res;
            object obj = objOrig;
            object[] args = AutoWrap.ПолучитьМассивРеальныхОбъектов(argsOrig);
            if (obj is AutoWrap Объект)
            {
                obj = Объект.O;

                if (Объект.ЭтоТип)
                {
                    res = AutoWrap.ОбернутьОбъект(Объект.T.InvokeMember(имяМетода, BindingFlags.DeclaredOnly |
                    BindingFlags.Public | BindingFlags.NonPublic |
                    BindingFlags.Static | BindingFlags.InvokeMethod, null, null, args));

                    AutoWrap.УстановитьИзмененияВМассиве(argsOrig, args);

                    return res;
                }
            }

            Type T = obj.GetType();
            res = AutoWrap.ОбернутьОбъект(T.InvokeMember(имяМетода, BindingFlags.DeclaredOnly |
              BindingFlags.Public | BindingFlags.NonPublic |
              BindingFlags.Instance | BindingFlags.InvokeMethod, null, obj, args));

            AutoWrap.УстановитьИзмененияВМассиве(argsOrig, args);
            return res;
        }
        public object MethodInfo_Invoke(object методИнфо, object объект, params object[] argsOrig)
        {
            object[] args = AutoWrap.ПолучитьМассивРеальныхОбъектов(argsOrig);
            var mf = (MethodInfo)(((AutoWrap)методИнфо).O);

            return AutoWrap.ОбернутьОбъект(mf.Invoke(ПолучитьРеальныйОбъект(объект), args));
        }
        public object Or(params object[] параметры1)
        {
            if (параметры1.Length == 0)
                return null;

            object[] параметры = AutoWrap.ПолучитьМассивРеальныхОбъектов(параметры1);
            var парам = параметры[0];
            var тип = парам.GetType();

            long res = (long)Convert.ChangeType(парам, typeof(long));

            for (int i = 1; i < параметры.Length; i++)
                res |= (long)Convert.ChangeType(параметры[i], typeof(long));


            if (тип.IsEnum)
            {
                var ТипЗначений = Enum.GetUnderlyingType(тип);
                var number = Convert.ChangeType(res, ТипЗначений);
                return AutoWrap.ОбернутьОбъект(Enum.ToObject(тип, number));
            }

            return AutoWrap.ОбернутьОбъект(Convert.ChangeType(res, тип));
        }
        public Object ПолучитьSafeArrayИзЭнумератора(Object массив)
        {
            List<object> res = new List<object>();

            var Enumerable = ((IEnumerable)(AutoWrap.ПолучитьРеальныйОбъект(массив))).GetEnumerator();

            while (Enumerable.MoveNext())
            {
                res.Add(AutoWrap.ОбернутьОбъект(Enumerable.Current));
            }

            return res.ToArray();
        }
        public Object ВМассив(Object массив)
        {
            return ПолучитьSafeArrayИзЭнумератора(массив);
        }
        public object ПолучитьМассивТипов()
        {
            var list = new List<object>
            {
                "System.String",
                AutoWrap.ОбернутьОбъект(DateTime.Now),
                AutoWrap.ОбернутьОбъект(true),
                AutoWrap.ОбернутьОбъект((Byte)45),
                AutoWrap.ОбернутьОбъект((Decimal)48.789),
                AutoWrap.ОбернутьОбъект(51.51),
                AutoWrap.ОбернутьОбъект((Single)11.11),
                AutoWrap.ОбернутьОбъект(11),
                AutoWrap.ОбернутьОбъект(789988778899),
                AutoWrap.ОбернутьОбъект((SByte)45),
                AutoWrap.ОбернутьОбъект((Int16)66),
                AutoWrap.ОбернутьОбъект((UInt32)77),
                AutoWrap.ОбернутьОбъект((UInt64)88888888888888),
                AutoWrap.ОбернутьОбъект((UInt16)102)
            };

            return new AutoWrap(list);
        }
        public object ОбновитьДанныеОметодахИСвойствах(object objOrig)
        {
            return AutoWrap.ОбернутьОбъект(AutoWrap.ПолучитьРеальныйОбъект(objOrig));
        }

        #endregion

        #region Private Methods

        private MethodInfo ПолучитьMethodInfoОберткиСобытий(string имяСвойства, object реальныйОбъект)
        {
            Type тип = реальныйОбъект.GetType();
            Type genType = typeof(КомВраперДляСобытий<>);
            Type constructed = genType.MakeGenericType(тип);

            // Now use reflection to invoke the method on constructed type
            PropertyInfo pi = constructed.GetProperty(имяСвойства);
            if (pi != null)
            {
                MethodInfo функция = (MethodInfo)pi.GetValue(null, null);

                return функция;
            }

            return null;
        }
        private Type ТипДляСоздатьОбъект(object типOrig)
        {
            object Тип = типOrig;
            string ТипСтр = Тип.ToString();
            if (Тип is AutoWrap)
            { Тип = ((AutoWrap)Тип).O; }
            else if (Тип is string тип)
            {
                Тип = НайтиТип(тип);
                if (Тип == null)
                {
                    string ошибка = " неверный тип " + ТипСтр;
                    MessageBox.Show(ошибка);
                    throw new COMException(ошибка);
                }
            }
            return (Type)Тип;
        }

        #endregion
    }
}
