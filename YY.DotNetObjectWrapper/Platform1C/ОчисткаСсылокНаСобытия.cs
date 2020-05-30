using System;
using System.Reflection;

namespace YY.DotNetObjectWrapper.Platform1C
{
    public static class ОчисткаСсылокНаСобытия
    {
        #region Public Methods

        public static void Очистить(object obj)
        {
            if (obj == null)
                return;

            var тип = obj.GetType();
            var Events = тип.GetEvents();

            foreach (var Event in Events)
            {

                FieldInfo fi = ПолучитьПолеРекурсивно(тип, Event.Name);

                if (fi == null)
                    continue;

                var value = (Delegate)fi.GetValue(obj);
                if (value != null)
                {
                    foreach (var eventDelegate in value.GetInvocationList())
                        Event.RemoveEventHandler(obj, eventDelegate);
                }
            }
        }

        #endregion

        #region Private Methods

        private static FieldInfo ПолучитьПолеРекурсивно(Type тип, string имяПоля)
        {
            if (тип == null)
                return null;

            FieldInfo fi = тип.GetField(имяПоля, BindingFlags.NonPublic | BindingFlags.Instance);

            if (fi == null)
                return ПолучитьПолеРекурсивно(тип.BaseType, имяПоля);

            return fi;
        }

        #endregion
    }
}
