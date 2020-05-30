using System.Runtime.InteropServices;

namespace YY.DotNetObjectWrapper.Platform1C
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class MyEnumerableClass
    {
        #region Private Members

        private readonly ТипизированныйЭнумератор _энумератор;

        #endregion

        #region Constructors

        public MyEnumerableClass(ТипизированныйЭнумератор энумератор)
        {
            this._энумератор = энумератор;

        }

        #endregion

        #region Public Methods

        [DispId(-4)]
        public object GetEnumerator()
        {
            return _энумератор;
        }

        #endregion
    }
}
