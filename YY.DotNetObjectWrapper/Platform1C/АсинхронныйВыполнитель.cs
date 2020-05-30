using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using YY.DotNetObjectWrapper.Service;

namespace YY.DotNetObjectWrapper.Platform1C
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("ECA3754A-0297-46F6-BA49-C23F204C8EF2")]
    [ComSourceInterfaces(typeof(IВрапперДляАсинхронныйВыполнитель))]
    public class АсинхронныйВыполнитель
    {
        #region Public Members

        public delegate void СобытиеОкончанияЗадачиDelgate(object задача, object данныеКЗадаче);
        public object Объект;        
        public event СобытиеОкончанияЗадачиDelgate ПриОкончанииВыполненияЗадачи;

        #endregion      

        #region Private Members

        private readonly SynchronizationContext _sc;

        #endregion

        #region Constructor

        public АсинхронныйВыполнитель()
        {
            if (SynchronizationContext.Current == null)
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());

            _sc = SynchronizationContext.Current;


        }

        #endregion

        #region Public Methods

        public void Выполнить(object value, object данныеКЗадаче)
        {
            var Задача = AutoWrap.ПолучитьРеальныйОбъект(value);

            if (this.ПриОкончанииВыполненияЗадачи != null) //Событие();
            {
                var типПараметра = Задача.GetType().GenericTypeArguments[0];

                //  ДляВыполненияЗадачи<dynamic>.Выполнить((dynamic)Задача, this);

                Type genType = typeof(ДляВыполненияЗадачи<>);
                Type constructed = genType.MakeGenericType(new Type[] { типПараметра });

                // Now use reflection to invoke the method on constructed type
                // var mi = constructed.GetMethod("Выполнить");
                constructed.InvokeMember("Выполнить", BindingFlags.DeclaredOnly |
                    BindingFlags.Public | BindingFlags.NonPublic |
                    BindingFlags.Static | BindingFlags.InvokeMethod, null, null, new[] { Задача, this, данныеКЗадаче });
            }

        }
        public void Оповестить(object задача, object данныеКЗадаче)
        {
            if(_sc != null && задача != null && данныеКЗадаче != null)
                _sc.Post(d => ПриОкончанииВыполненияЗадачи(AutoWrap.ОбернутьОбъект(задача), данныеКЗадаче), null);
        }

        #endregion
    }
}
