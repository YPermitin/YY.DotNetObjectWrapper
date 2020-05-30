using System.Runtime.InteropServices;

namespace YY.DotNetObjectWrapper.Platform1C
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IEventFor1C
    {
        [DispId(0x60020000)]
        void Событие();

        [DispId(0x60020001)]
        void СобытиеСПараметром(object value);
    }    
}
