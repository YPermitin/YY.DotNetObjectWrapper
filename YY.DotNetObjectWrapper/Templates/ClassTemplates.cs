namespace YY.DotNetObjectWrapper.Templates
{    
    public static class ClassTemplates
    {
        public static string Template =
        #region TemplateFor1C8
@"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;


[ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(""{0}"")]
    public interface IВрапперДля{1}
        {{
     
            [DispId(0x00000001)]
            void ОшибкаСобытия(String событие, object value, object исключение);
            {2}
        }}


        [ComVisible(true)]
        [ClassInterface(ClassInterfaceType.AutoDual)]
        [Guid(""{3}"")]
        [ComSourceInterfaces(typeof(IВрапперДля{1}))]
        public class ВрапперДля{1}
        {{

        [ComVisible(false)]
        public delegate void Событие_Delgate();
        [ComVisible(false)]
        public delegate void СобытиеСПараметром_Delgate(object value);
        [ComVisible(false)]
        public delegate void СобытиеСПараметрами_Delgate(String событие, object value, object исключение);

            public {4} РеальныйОбъект;
            dynamic AutoWrap;
            private SynchronizationContext Sc;
            public event СобытиеСПараметрами_Delgate ОшибкаСобытия;
           
            {5}
            private Object thisLock = new Object();
          

            void ОтослатьСобытие(Событие_Delgate Событие, string ИмяСобытия)
            {{
                

                Task.Run(() =>
                {{
                    if (Событие != null) //Событие();
                    {{
                        lock (thisLock)
                        {{

                            try
                            {{
                                Sc.Send(d => Событие(), null);
                            }}

                            catch (Exception ex)
                            {{
                                try
                                {{
                                  
                                     Sc.Send(d => ОшибкаСобытия(ИмяСобытия, null, AutoWrap.ОбернутьОбъект(ex)), null);
                                }}
                                catch (Exception)
                                {{
                                }}

                            }}
                        }}
                    }}
                }});
            }}

            void ОтослатьСобытиеСПараметром(СобытиеСПараметром_Delgate Событие, object value, string ИмяСобытия)
            {{
                Task.Run(() =>
                {{
                   

                    if (Событие != null) //Событие();
                    {{
                        lock (thisLock)
                        {{
                            try
                            {{
                                Sc.Send(d => Событие(AutoWrap.ОбернутьОбъект(value)), null);
                            }}
                            catch (Exception ex)
                            {{
                                try
                                {{
                                    
                                    Sc.Send(d => ОшибкаСобытия(ИмяСобытия, AutoWrap.ОбернутьОбъект(value), AutoWrap.ОбернутьОбъект(ex)), null);
                                }}
                                catch (Exception)
                                {{
                                }}
                            }}

                        }}
                    }}
                }});


            }}
            public ВрапперДля{1}(object AutoWrap,{4}  РеальныйОбъект)
            {{
              
                if (SynchronizationContext.Current==null)
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());

                Sc = SynchronizationContext.Current;


                this.РеальныйОбъект = РеальныйОбъект;
                this.AutoWrap=AutoWrap;
                {6}
           }}

public static object СоздатьОбъект(object AutoWrap,{4}  РеальныйОбъект)
            {{
              
                return new ВрапперДля{1}(AutoWrap,РеальныйОбъект);
           }}
}}";
        #endregion
        public static string Template77 =
        #region Template1C77
@"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;

 [Guid(""ab634004-f13d-11d0-a459-004095e1daea""), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAsyncEvent
        {{
            void SetEventBufferDepth(int lDepth);
            void GetEventBufferDepth(ref int plDepth);
            void ExternalEvent(string bstrSource, string bstrMessage, string bstrData);
            void CleanBuffer();
        }}

        public class ВрапперДля{0}77
    {{
        {1} РеальныйОбъект;
        {2}

       
        private SynchronizationContext Sc;
        IAsyncEvent СобытиеДля1С;
        private Object thisLock = new Object();
        public object ПоследняяОшибка;
        public   ВрапперДля{0}77(Object ГлобальныйКонтекст, {1} РеальныйОбъект)
          {{

            СобытиеДля1С = ГлобальныйКонтекст as IAsyncEvent;
            if (SynchronizationContext.Current == null)
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());

            Sc = SynchronizationContext.Current;


            this.РеальныйОбъект = РеальныйОбъект;
            
            {3}

        }}

        void ОтослатьСобытие(string ИмяСобытия)
{{


    Task.Run(() =>
    {{

        lock (thisLock)
        {{

            try
            {{
                Sc.Send(d => СобытиеДля1С.ExternalEvent(""{0}"", ИмяСобытия, """"), null);
            }}

            catch (Exception ex)
            {{
                try
                {{
                    ПоследняяОшибка = new {{ Событие = ИмяСобытия, Исключение = ex }};
                    Sc.Send(d => СобытиеДля1С.ExternalEvent(""{0}"", ""ОшибкаСобытия"", """"), null);
                }}
                catch (Exception)
                {{
                }}

            }}
        }}

    }});
}}

public static object СоздатьОбъект(Object ГлобальныйКонтекст, {1} РеальныйОбъект)
{{

    return new ВрапперДля{0}77(ГлобальныйКонтекст, РеальныйОбъект);
}}
    }}
";
        #endregion
    }
}
