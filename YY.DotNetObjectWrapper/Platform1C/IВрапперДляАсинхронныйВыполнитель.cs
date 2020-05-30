using System.Runtime.InteropServices;

namespace YY.DotNetObjectWrapper.Platform1C
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid("FCD122A7-0913-409E-AB96-2AD5BFA148EC")]
    public interface IВрапперДляАсинхронныйВыполнитель
    {
        [DispId(0x00000001)]
        void ПриОкончанииВыполненияЗадачи(object задача, object данныеКЗадаче);
    }
}
