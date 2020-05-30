using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using YY.DotNetObjectWrapper.Platform1C;

namespace YY.DotNetObjectWrapper.Service
{
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ComVisible(true)]
    [Guid("1521B4B2-F38B-4CAD-BB45-3C6E1F00982F")]
    public class AutoWrap : IReflect
    {
        #region Static Constructors

        static AutoWrap()
        {
            МетодыObject = new Dictionary<string, bool>();
            foreach (var метод in typeof(object).GetMethods(BindingFlags.Public | BindingFlags.Instance))
                МетодыObject[метод.Name] = true;
        }

        #endregion

        #region Public Static Members

        public static bool ЭтоСемерка;
        public static bool ВыводитьСообщениеОбОшибке = true;
        public static Exception ПоследняяОшибка;
        public static Dictionary<string, bool> МетодыObject;

        #endregion

        #region Private Static Members

        private static BindingFlags _staticBinding = BindingFlags.Public | BindingFlags.Static;

        #endregion

        #region Public Static Methods

        public static object ПолучитьРеальныйОбъект(object obj)
        {
            if (obj is AutoWrap)            
                return ((AutoWrap)obj).O;            

            return obj;
        }
        public static object ОбернутьОбъект(object obj)
        {
            if (obj != null)
            {
                Type Тип = obj.GetType();

                if (Тип == typeof(IntPtr) || Тип == typeof(UIntPtr))
                    return new AutoWrap(obj);

                if (Тип == typeof(Char))
                    return Convert.ToString(((char)obj));

                if (ЭтоСемерка)
                {

                    if (Тип == typeof(Decimal)) return ((Decimal)obj).ToString(CultureInfo.InvariantCulture);
                    if (Тип.IsPrimitive)
                    {
                        if ((Тип == typeof(Int64) || Тип == typeof(UInt32) || Тип == typeof(UInt64) || Тип == typeof(UInt16) || Тип == typeof(SByte)))
                            obj = Convert.ChangeType(obj, typeof(string), CultureInfo.InvariantCulture);
                    }
                    else if (!(Тип == typeof(DateTime)
                               || Тип == typeof(String)
                               || Тип == typeof(Decimal)
                               || Тип.IsCOMObject)
                             )
                        obj = new AutoWrap(obj);
                }
                else
                {
                    if (Тип.IsArray)
                    {
                        Type ТипМассива = Тип.GetElementType();
                        if (ТипМассива != null)
                            Тип = ТипМассива;
                    }

                    if (!(Тип.IsPrimitive
                    || Тип == typeof(Decimal)
                    || Тип == typeof(DateTime)
                    || Тип == typeof(String)
                    || Тип.IsCOMObject)
                          )
                        obj = new AutoWrap(obj);
                }
            }
            return obj;
        }

        #endregion

        #region Private Static Methods

        private static Exception NotImplemented()
        {
            var method = new StackTrace(true).GetFrame(1).GetMethod().Name;
            Debug.Assert(false, method);
            return new NotImplementedException(method);
        }
        internal static object[] ПолучитьМассивРеальныхОбъектов(object[] args)
        {
            if (args.Length == 0)
                return args;

            object[] res = new object[args.Length];
            for (int i = 0; i < args.Length; i++)
            {
                res[i] = ПолучитьРеальныйОбъект(args[i]);
            }

            return res;
        }
        internal static void УстановитьИзмененияВМассиве(object[] оригинальныйМассив, object[] массивРеальныхОбъектов)
        {
            if (оригинальныйМассив == массивРеальныхОбъектов)
                return;

            for (int i = 0; i < оригинальныйМассив.Length; i++)
            {
                object obj = оригинальныйМассив[i];
                if (obj is AutoWrap)
                {
                    AutoWrap элемент = (AutoWrap)obj;
                    if (!Equals(элемент.O, массивРеальныхОбъектов[i]))
                        оригинальныйМассив[i] = ОбернутьОбъект(массивРеальныхОбъектов[i]);
                }
                else
                {
                    if (!Equals(оригинальныйМассив[i], массивРеальныхОбъектов[i]))
                        оригинальныйМассив[i] = ОбернутьОбъект(массивРеальныхОбъектов[i]);
                }
            }
        }

        #endregion

        #region Public Properties

        public Type UnderlyingSystemType
        {
            get { return T.UnderlyingSystemType; }
        }

        #endregion

        #region Private Members

        protected internal object O;
        protected internal Type T;

        internal FieldInfo[] Поля;
        internal MemberInfo[] Мемберы;
        internal MethodInfo[] Методы;
        internal PropertyInfo[] Свойства;
        
        internal bool ЭтоТип;
        internal bool ЭтоExpandoObject;

        #endregion

        #region Constructors

        public AutoWrap(object obj)
        {
            O = obj;
            if (O is Type)
            {
                T = O as Type;
                ЭтоТип = true;
            }
            else
            {
                T = O.GetType();
                ЭтоТип = false;
                ЭтоExpandoObject = O is ExpandoObject;
                ДанныеДляТипа.Инициализировать(this);
            }
        }
        public AutoWrap(object obj, Type type)
        {
            O = obj;
            T = type;
            ЭтоТип = false;
        }

        #endregion

        #region Public Methods

        public FieldInfo[] GetFields(BindingFlags bindingAttr)
        {
            if (Поля != null) return Поля;

            if (ЭтоТип)
                Поля = T.GetFields(_staticBinding);
            else if (ЭтоExpandoObject)
            {
                Поля = (from KeyValue in (IEnumerable<KeyValuePair<string, object>>)O
                        where (!(KeyValue.Value is Delegate))
                        select new DynamicFieldInfo(KeyValue.Key)).ToArray();
            }
            else
                Поля = T.GetFields(bindingAttr);

            return Поля;
        }
        public MemberInfo[] GetMember(string name, BindingFlags bindingAttr)
        {
            if (ЭтоТип)
                return T.GetMember(name, _staticBinding);
            if (ЭтоExpandoObject)
                throw NotImplemented();
            return T.GetMember(name, bindingAttr);
        }
        public MemberInfo[] GetMembers(BindingFlags bindingAttr)
        {
            if (Мемберы != null) return Мемберы;

            if (ЭтоТип)
                Мемберы = T.GetMembers(_staticBinding);
            else if (ЭтоExpandoObject)
                throw NotImplemented();
            else
                Мемберы = T.GetMembers(bindingAttr);

            return Мемберы;
        }
        public MethodInfo GetMethod(string name, BindingFlags bindingAttr)
        {
            if (ЭтоТип)
                return T.GetMethod(name, _staticBinding);
            if (ЭтоExpandoObject)
                throw NotImplemented();
            return T.GetMethod(name, bindingAttr);
        }
        public MethodInfo GetMethod(string name, BindingFlags bindingAttr, Binder binder, Type[] types, ParameterModifier[] modifiers)
        {
            if (ЭтоТип)
                return T.GetMethod(name, _staticBinding, binder, types, modifiers);
            if (ЭтоExpandoObject)
                throw NotImplemented();
            return T.GetMethod(name, bindingAttr, binder, types, modifiers);
        }
        public MethodInfo[] GetMethods(BindingFlags bindingAttr)
        {
            if (Методы != null) return Методы;
            if (ЭтоТип)
                Методы = T.GetMethods(_staticBinding);
            else if (ЭтоExpandoObject)
            {
                Методы = (from KeyValue in (IEnumerable<KeyValuePair<string, object>>)O
                          where (KeyValue.Value is Delegate)
                          select new DynamicMethodInfo(KeyValue.Key, ((Delegate)KeyValue.Value).Method))
                 .Concat(from y in typeof(object).GetMethods(bindingAttr)
                         select y).ToArray();
            }
            else
                Методы = T.GetMethods(bindingAttr);

            return Методы;
        }
        public PropertyInfo[] GetProperties(BindingFlags bindingAttr)
        {
            if (Свойства != null) return Свойства;

            if (ЭтоТип)
                Свойства = T.GetProperties(_staticBinding);
            else if (ЭтоExpandoObject)
                return new PropertyInfo[0];
            else
                Свойства = T.GetProperties(bindingAttr);

            return Свойства;
        }
        public PropertyInfo GetProperty(string name, BindingFlags bindingAttr, Binder binder, Type returnType, Type[] types, ParameterModifier[] modifiers)
        {
            if (ЭтоТип)
                return T.GetProperty(name, _staticBinding, binder, returnType, types, modifiers);
            if (ЭтоExpandoObject)
                throw NotImplemented();
            return T.GetProperty(name, bindingAttr, binder, returnType, types, modifiers);
        }
        public PropertyInfo GetProperty(string name, BindingFlags bindingAttr)
        {
            if (ЭтоТип)
                return T.GetProperty(name, _staticBinding);
            if (ЭтоExpandoObject)
                throw NotImplemented();
            return T.GetProperty(name, bindingAttr);
        }

        #region IReflect Members

        public FieldInfo GetField(string name, BindingFlags bindingAttr)
        {
            if (ЭтоТип)
                return T.GetField(name, _staticBinding);
            if (ЭтоExpandoObject)
                throw NotImplemented();
            return T.GetField(name, bindingAttr);
        }
        public void ПроверитьНаДоступКПолям(ref BindingFlags invokeAttr, int количествоПараметров)
        {
            if ((
                (
                (invokeAttr & BindingFlags.PutDispProperty) == BindingFlags.PutDispProperty)
                || (invokeAttr.HasFlag(BindingFlags.PutRefDispProperty))
                )
                && (количествоПараметров == 1))
            {
                invokeAttr = invokeAttr | BindingFlags.SetField;

            }
            else if (((invokeAttr & BindingFlags.GetProperty) == BindingFlags.GetProperty) && (количествоПараметров == 0))
            {
                invokeAttr = invokeAttr | BindingFlags.GetField;

            }
        }
        public object InvokeMemberExpandoObject(string name, BindingFlags invokeAttr, object[] args)
        {
            if (invokeAttr.HasFlag(BindingFlags.InvokeMethod))
            {

                return ((Delegate)((IDictionary<string, object>)((ExpandoObject)O))[name]).DynamicInvoke(args);
            }

            if (invokeAttr.HasFlag(BindingFlags.GetField))
                return ((IDictionary<string, object>)((ExpandoObject)O))[name];

            if (invokeAttr.HasFlag(BindingFlags.SetField))
                ((IDictionary<string, object>)((ExpandoObject)O))[name] = args[0];

            return null;
        }
        public object InvokeMember(string name, BindingFlags invokeAttr, Binder binder, object target, object[] argsOrig, ParameterModifier[] modifiers, CultureInfo culture, string[] namedParameters)
        {        
            // Unwrap any AutoWrap'd objects (they need to be raw if a paramater)
            if (name == "[DISPID=-4]")
            {
                IEnumVARIANT rez = new EnumVariantImpl(((IEnumerable)O).GetEnumerator());
                return rez;
            }

            object[] args = ПолучитьМассивРеальныхОбъектов(argsOrig);
            if(culture == null)
                culture = CultureInfo.InvariantCulture;

            object obj;

            try
            {

                if (T.IsEnum && !((invokeAttr & BindingFlags.InvokeMethod) == BindingFlags.InvokeMethod))
                    return ОбернутьОбъект(Enum.Parse(T, name));

                ПроверитьНаДоступКПолям(ref invokeAttr, args.Length);

                if (ЭтоТип)
                    obj = T.InvokeMember(name, invokeAttr, binder, null, args, modifiers, culture, namedParameters);
                else if (ЭтоExpandoObject)
                {
                    if (invokeAttr.HasFlag(BindingFlags.InvokeMethod) && МетодыObject.ContainsKey(name))
                        obj = T.InvokeMember(name, invokeAttr, binder, O, args, modifiers, culture, namedParameters);
                    else
                        obj = InvokeMemberExpandoObject(name, invokeAttr, args);
                }
                else
                    obj = T.InvokeMember(name, invokeAttr, binder, O, args, modifiers, culture, namedParameters);
            }
            catch (Exception e)
            {
                ПоследняяОшибка = e;
                string Ошибка = "Ошибка в методе " + name + " " + e.Message + " " + e.Source;

                if (e.InnerException != null)
                    Ошибка = Ошибка + "\r\n" + e.InnerException;

                if (ВыводитьСообщениеОбОшибке)
                {
                    MessageBox.Show(Ошибка);
                    MessageBox.Show(e.StackTrace);
                    MessageBox.Show(invokeAttr.ToString());
                }
                throw new COMException(Ошибка);
            }

            // Так как параметры могут изменяться (OUT) и передаются по ссылке
            // нужно обратно обернуть параметры
            УстановитьИзмененияВМассиве(argsOrig, args);

            return ОбернутьОбъект(obj);
        }

        #endregion

        #endregion

        #region Private Methods

        private void ДобавитьМетодыObject()
        {

        }

        #endregion
    }
}
