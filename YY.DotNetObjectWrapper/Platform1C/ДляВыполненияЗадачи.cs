namespace YY.DotNetObjectWrapper.Platform1C
{
    public class ДляВыполненияЗадачи<TResult>
    {
        static public void Выполнить(System.Threading.Tasks.Task<TResult> задача, АсинхронныйВыполнитель выполнитель, object данныеКЗадаче)
        {
            задача.ContinueWith(t => {
                выполнитель.Оповестить(t, данныеКЗадаче);
            });
        }
    }
}
