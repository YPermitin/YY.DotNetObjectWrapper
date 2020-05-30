using System;
using System.Runtime.InteropServices;

namespace YY.DotNetObjectWrapper.AddIn
{
    [ComVisible(true)]
    [ProgId("AddIn.GlobalContext1C")]
    [Guid("8693BBEC-C964-4478-AFCB-E8D15FD8F4F6")]
    public class GlobalContext1C : IInitDone, ILanguageExtender
    {
        #region Private Member Variables

        private Object _глобальныйКонтекст;

        #endregion
        
        #region Public Methods

        public void RegisterExtensionAs([MarshalAs(UnmanagedType.BStr)] ref string extensionName)
        {
            extensionName = "GlobalContext1C";
        }
        public void Init([MarshalAs(UnmanagedType.IDispatch)] object connection)
        {
            _глобальныйКонтекст = connection;
            Marshal.GetIUnknownForObject(_глобальныйКонтекст);

        }
        public void GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] info)
        {

            info[0] = 1000;

        }
        public void Done()
        {
            Marshal.Release(Marshal.GetIDispatchForObject(_глобальныйКонтекст));
            Marshal.ReleaseComObject(_глобальныйКонтекст);
            _глобальныйКонтекст = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        public void FindMethod([MarshalAs(UnmanagedType.BStr)] String methodName, ref Int32 methodNum)
        {
            throw new COMException("Нет методов");
        }
        public void GetNParams(Int32 methodNum, ref Int32 pParams)
        {
            throw new COMException("Нет методов и параметров");
        }
        public void HasRetVal(Int32 methodNum, ref Boolean retValue)
        {
            throw new COMException("Нет методов и параметров");
        }
        public void CallAsFunc(Int32 methodNum, [MarshalAs(UnmanagedType.Struct)] ref object retValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {
            throw new COMException("Нет методов и параметров");
        }
        public void GetNProps(ref Int32 props)
        {

        }
        public void FindProp([MarshalAs(UnmanagedType.BStr)] String propName, ref Int32 propNum)
        {
            if (propName.Trim().ToUpper() == "ГЛОБАЛЬНЫЙКОНТЕКСТ")
            {
                propNum = 0;
                return;
            }
            throw new COMException(String.Format("Свойство {0} не поддерживается", propName));

        }
        public void GetPropName(Int32 propNum, Int32 propAlias, [MarshalAs(UnmanagedType.BStr)] ref String propName)
        {

        }
        public void GetPropVal(Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
        {
            if (propNum == 0) propVal = _глобальныйКонтекст;
        }
        public void SetPropVal(Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
        {

        }
        public void IsPropReadable(Int32 propNum, ref bool propRead)
        {
            propRead = true;
        }
        public void IsPropWritable(Int32 propNum, ref Boolean propWrite)
        {
            propWrite = false;
        }
        public void GetNMethods(ref Int32 pMethods)
        {

        }
        public void GetMethodName(Int32 methodNum, Int32 methodAlias, [MarshalAs(UnmanagedType.BStr)] ref String methodName)
        {

        }
        public void GetParamDefValue(Int32 methodNum, Int32 paramNum, [MarshalAs(UnmanagedType.Struct)]ref object paramDefValue)
        {

        }
        public void CallAsProc(int lMethodNum, [MarshalAs(UnmanagedType.SafeArray)] ref Array paParams)
        {

        }
        public void CallAsFunc(int lMethodNum, ref object pvarRetValue, [MarshalAs(UnmanagedType.SafeArray)] ref Array paParam)
        {

        }

        #endregion
    }
}
