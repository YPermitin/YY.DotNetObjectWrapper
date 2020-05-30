using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Globalization;
using System.Reflection;

namespace YY.DotNetObjectWrapper.Service
{
    public class DynamicFieldInfo : FieldInfo
    {        
        #region Constructor

        public DynamicFieldInfo(string имяПоля)
        {
            _name = имяПоля;
        }

        #endregion

        #region Private Members

        private string _name;

        #endregion

        #region Public Properties

        /// <summary>Получает дескриптор представления внутренних метаданных этого поля.</summary>
        /// <returns>Дескриптор представления внутренних метаданных этого поля.</returns>
        public override RuntimeFieldHandle FieldHandle
        {
            get
            {
                return new RuntimeFieldHandle();
            }
        }
        /// <summary>Получает атрибуты, связанные с этим полем.</summary>
        /// <returns>
        ///   <see cref="F:System.Reflection.FieldAttributes.PrivateScope" />
        /// </returns>
        public override FieldAttributes Attributes
        {
            get
            {
                return FieldAttributes.PrivateScope;
            }
        }
        /// <summary>Получает тип, объявляющий данное поле.</summary>
        /// <returns>Значение NULL во всех случаях.</returns>
        public override Type DeclaringType
        {
            get
            {
                return null;
            }
        }
        /// <summary>Возвращает тип данного поля.</summary>
        /// <returns>Тип <see cref="T:System.Object" />.</returns>
        public override Type FieldType
        {
            get
            {
                return typeof(Object);
            }
        }
        /// <summary>Получает тип члена, которым является данное поле.Указывает тип класса, производного от <see cref="T:System.Reflection.MemberInfo" />, который наследуется текущим классом.</summary>
        /// <returns>
        ///   <see cref="F:System.Reflection.MemberTypes.Field" />, because this class derives from <see cref="T:System.Reflection.FieldInfo" />.</returns>
        public override MemberTypes MemberType
        {
            get
            {
                return MemberTypes.Field;
            }
        }
        /// <summary>Получает имя данного поля.</summary>
        /// <returns>Имя этого поля.</returns>
        public override string Name
        {
            get
            {
                return _name;
            }
        }
        /// <summary>Получает объект класса, который использовался для извлечения данного экземпляра путем отражения.</summary>
        /// <returns>Тип, которым объявлен данный метод.</returns>
        public override Type ReflectedType
        {
            get
            {
                return DeclaringType;
            }
        }

        #endregion

        #region Public Methods

        /// <summary>Возвращает массив, содержащий настраиваемые атрибуты, вложенные в это поле, выполняя поиск только атрибутов заданного типа.</summary>
        /// <returns>Массив объектов <see cref="T:System.Reflection.FieldInfo" />, содержащий нулевые элементы.</returns>
        /// <param name="t">Искомый тип атрибута.</param>
        /// <param name="inherit">Значение true, чтобы выполнить поиск в иерархии наследования этого члена для поиска атрибутов.</param>
        public override object[] GetCustomAttributes(Type t, bool inherit)
        {
            return new FieldInfo[0];
        }
        /// <summary>Возвращает массив, содержащий настраиваемые атрибуты, вложенные в это поле.</summary>
        /// <returns>Массив объектов <see cref="T:System.Reflection.FieldInfo" />, содержащий нулевые элементы.</returns>
        /// <param name="inherit">Значение true, чтобы выполнить поиск в иерархии наследования этого члена для поиска атрибутов.</param>
        public override object[] GetCustomAttributes(bool inherit)
        {
            return new FieldInfo[0];
        }
        /// <summary>Определяет, добавлен ли в это поле указанный тип атрибута.</summary>
        /// <returns>Значение false во всех случаях.</returns>
        /// <param name="type">Искомый тип атрибута.</param>
        /// <param name="inherit">Значение true, чтобы выполнить поиск в иерархии наследования этого члена для поиска атрибутов.</param>
        public override bool IsDefined(Type type, bool inherit)
        {
            return false;
        }
        /// <summary>Задает значение поля, используя заданное значение, язык и региональные параметры, а также сведения о привязке.</summary>
        /// <param name="obj">Объект, для которого будет установлено значение поля.</param>
        /// <param name="value">Значение, присваиваемое полю.</param>
        /// <param name="invokeAttr">Побитовое сочетание значений перечисления, которое используется для управления привязкой.</param>
        /// <param name="binder">Набор свойств, который позволяет осуществлять связывание, приведение типов аргументов и вызов элементов с помощью отражения.</param>
        /// <param name="culture">Предоставляет сведения, касающиеся языка и региональных параметров или языкового стандарта.Используется для надлежащего форматирования чисел, дат и строк.</param>
        public override void SetValue(object obj, object value, BindingFlags invokeAttr, Binder binder, CultureInfo culture)
        {
            ((IDictionary<string, object>)((ExpandoObject)obj))[_name] = value;
        }
        /// <summary>Получает значение поля.</summary>
		/// <returns>Значение поля.</returns>
		/// <param name="obj">Объект, значение поля которого будет возвращено.</param>
		public override object GetValue(object obj)
        {
            return ((IDictionary<string, object>)((ExpandoObject)obj))[_name];
        }

        #endregion
    }
}
