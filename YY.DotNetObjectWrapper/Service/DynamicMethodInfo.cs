using System;
using System.Globalization;
using System.Reflection;

namespace YY.DotNetObjectWrapper.Service
{
    public class DynamicMethodInfo : MethodInfo
    {
        #region Public Properties

        public override Type DeclaringType
        {
            get { return _mi.DeclaringType; }
        }
        public override string Name
        {
            get { return _name; }
        }
        public override Type ReflectedType
        {
            get { return _mi.ReflectedType; }
        }

        #endregion

        #region Private Members

        string _name;
        MethodInfo _mi;

        #endregion

        #region Constructor

        public DynamicMethodInfo(string name, MethodInfo mi)
        {
            _name = name;
            _mi = mi;
        }

        #endregion

        #region Public Methods

        public override MethodInfo GetBaseDefinition()
        {
            return _mi.GetBaseDefinition();
        }
        public override ICustomAttributeProvider ReturnTypeCustomAttributes
        {
            get { return _mi.ReturnTypeCustomAttributes; }
        }
        public override MethodAttributes Attributes
        {
            get { return _mi.Attributes; }
        }
        public override MethodImplAttributes GetMethodImplementationFlags()
        {
            return _mi.GetMethodImplementationFlags();
        }
        public override ParameterInfo[] GetParameters()
        {
            return _mi.GetParameters();
        }
        public override object Invoke(object obj, BindingFlags invokeAttr, Binder binder, object[] parameters, CultureInfo culture)
        {
            return _mi.Invoke(obj, invokeAttr, binder, parameters, culture);
        }
        public override RuntimeMethodHandle MethodHandle
        {
            get { return _mi.MethodHandle; }
        }
        public override object[] GetCustomAttributes(Type attributeType, bool inherit)
        {
            return _mi.GetCustomAttributes(attributeType, inherit);
        }
        public override object[] GetCustomAttributes(bool inherit)
        {
            return _mi.GetCustomAttributes(inherit);
        }
        public override bool IsDefined(Type attributeType, bool inherit)
        {
            return _mi.IsDefined(attributeType, inherit);
        }

        #endregion
    }
}
