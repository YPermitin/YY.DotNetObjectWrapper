using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using YY.DotNetObjectWrapper.Service;

namespace YY.DotNetObjectWrapper.Platform1C
{
    public class ТипизированныйЭнумератор : IEnumVARIANT
    {
        #region Private Members

        private const int SOk = 0;
        private const int SFalse = 1;
        private readonly System.Collections.IEnumerator _enumerator;
        private readonly Type _type;
        private readonly ДанныеДляТипа _данныеДляТипа;

        #endregion

        #region Constructors

        public ТипизированныйЭнумератор(System.Collections.IEnumerator enumerator, Type тип)
        {
            this._enumerator = enumerator;
            _type = тип;


            _данныеДляТипа = ДанныеДляТипа.ПолучитьДанныеДляТипа(_type);
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
                var current = _enumerator.Current;
                AutoWrap res = null;
                if (!((current == null) || !_type.IsAssignableFrom(current.GetType())))
                {
                    res = new AutoWrap(current, _type);
                    ДанныеДляТипа.ПрописатьПоля(res, _данныеДляТипа);
                }

                rgVar[0] = res;
                if (pceltFetched != IntPtr.Zero)
                    Marshal.WriteInt32(pceltFetched, 1);

                return SOk;
            }
            else
            {
                return SFalse;
            }
        }

        #endregion
    }
}
