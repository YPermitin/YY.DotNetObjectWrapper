using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace YY.DotNetObjectWrapper.Service
{
    public class EnumVariantImpl : IEnumVARIANT
    {
        #region Private Members

        private const int SOk = 0;
        private const int SFalse = 1;
        private readonly IEnumerator _enumerator;

        #endregion

        #region Constructor

        public EnumVariantImpl(IEnumerator enumerator)
        {
            _enumerator = enumerator;
        }

        #endregion

        #region Public Methods

        public IEnumVARIANT Clone()
        {
            throw new NotImplementedException();
        }
        public int Reset()
        {
            try
            {
                _enumerator.Reset();
            }
            catch
            {

                return SFalse;
            }

            return SOk;
        }
        public int Skip(int celt)
        {
            for (; celt > 0; celt--)
                if (!_enumerator.MoveNext())
                    return SFalse;
            return SOk;
        }
        public int Next(int celt, object[] rgVar, IntPtr pceltFetched)
        {
            if (celt == 1 && _enumerator.MoveNext())
            {
                rgVar[0] = AutoWrap.ОбернутьОбъект(_enumerator.Current);
                if (pceltFetched != IntPtr.Zero)
                    Marshal.WriteInt32(pceltFetched, 1);

                return SOk;
            }

            return SFalse;
        }

        #endregion
    }
}
