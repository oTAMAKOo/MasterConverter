
using System.Linq;

namespace MasterConverter
{
    public interface IContainerClass
    {
        void SetRecords(object[] values);
    }
    
    public class ContainerClass<T> : IContainerClass
    {
        public T[] records = null;

        public void SetRecords(object[] values)
        {
            records = values.Cast<T>().ToArray();
        }
    }
}
