using System;
using Extensions;

namespace MasterConverter
{
    public sealed class IndexData
    {
        public string[] records = null;
    }


    [Serializable]
    public sealed class RecordData
    {
        public string recordName = null;

        public RecordValue[] values = null;

        public ExcelCell[] cells = null;
    }
    
    [Serializable]
    public sealed class RecordValue
    {
        public string fieldName = null;

        public object value = null;
    }

    [Serializable]
    public sealed class ExcelCell : Extensions.ExcelCell
    {
        public int column = 0;
    }
}
