using System;
using System.Collections.Generic;
using System.Reflection;
using YY.DotNetObjectWrapper.Service;

namespace YY.DotNetObjectWrapper.Platform1C
{
    internal class ДанныеДляТипа
    {
        #region Public Static Members

        public static readonly BindingFlags BindingAttr = BindingFlags.Public | BindingFlags.Instance;

        #endregion

        #region Private Static Methods

        internal static ДанныеДляТипа ПолучитьДанныеДляТипа(Type T)
        {
            if (!СловарьТипов.TryGetValue(T, out var res))
            {
                res = new ДанныеДляТипа(T);
                СловарьТипов.Add(T, res);
            }

            return res;
        }
        internal static void ПрописатьПоля(AutoWrap объект, ДанныеДляТипа res)
        {

            объект.Поля = res._поля;
            объект.Мемберы = res._мемберы;
            объект.Методы = res._методы;
            объект.Свойства = res._свойства;
        }
        internal static void Инициализировать(AutoWrap объект)
        {
            if (объект.ЭтоТип || объект.ЭтоExpandoObject)
                return;

            ДанныеДляТипа res = ПолучитьДанныеДляТипа(объект.T);
            ПрописатьПоля(объект, res);
        }

        #endregion

        #region Private Static Members

        internal static Dictionary<Type, ДанныеДляТипа> СловарьТипов = new Dictionary<Type, ДанныеДляТипа>();

        #endregion

        #region Constructors

        internal ДанныеДляТипа(Type T)
        {
            _поля = T.GetFields(BindingAttr);
            _мемберы = T.GetMembers(BindingAttr);
            _методы = T.GetMethods(BindingAttr);
            _свойства = T.GetProperties(BindingAttr);
        }

        #endregion

        #region Private Members

        private readonly FieldInfo[] _поля;
        private readonly MemberInfo[] _мемберы;
        private readonly MethodInfo[] _методы;
        private readonly PropertyInfo[] _свойства;        

        #endregion        
    }
}
