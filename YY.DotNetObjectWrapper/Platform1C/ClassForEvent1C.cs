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
    [Guid("62F8156C-13B9-4484-B152-82023243E1D3")]
    [ComSourceInterfaces(typeof(IEventFor1C))]
    public class ClassForEvent1C
    {
        #region Public Properties

        [ComVisible(false)]
        public delegate void СобытиеDelgate();
        public delegate void СобытиеСПараметромDelgate(object value);
        public event СобытиеDelgate Событие;
        public event СобытиеСПараметромDelgate СобытиеСПараметром;
        public object Объект;

        #endregion

        #region Private Members

        private readonly SynchronizationContext _sc;

        #endregion

        #region Constructor

        public ClassForEvent1C(object объект, String событиеОбъекта, bool естьПараметр = false)
        {
            this.Объект = объект;
            BindingFlags bf = BindingFlags.DeclaredOnly | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;
            EventInfo ei = объект.GetType().GetEvent(событиеОбъекта, bf);
            if (!естьПараметр)
            {
                if (ei != null) ei.AddEventHandler(объект, new Action(ВнешнееСобытие));
            }
            else if (ei != null) ei.AddEventHandler(объект, new Action<object>(ВнешнееСобытиеСПараметром));

            if (SynchronizationContext.Current == null)
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());

            _sc = SynchronizationContext.Current;

        }

        #endregion

        #region Private Methods

        private void ВнешнееСобытие()
        {
            if (Событие != null && _sc != null)
                _sc.Send(d => Событие(), null);


        }
        private void ВнешнееСобытиеСПараметром(object value)
        {
            if (СобытиеСПараметром != null && _sc != null && value != null)
                _sc.Send(d => СобытиеСПараметром(AutoWrap.ОбернутьОбъект(value)), null);
        }

        #endregion
    }
}
