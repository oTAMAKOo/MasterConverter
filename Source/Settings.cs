
using System.IO;
using System.Reflection;

namespace MasterConverter
{
    public sealed class Settings
    {
        //----- params -----
        
        public class MasterSettings
        {
            /// <summary> マスター：タグ定義行 </summary>
            public int tagRow = 2;
            /// <summary> マスター：型名定義名行 </summary>
            public int dataTypeRow = 3;
            /// <summary> マスター：フィールド名定義行 </summary>
            public int fieldNameRow = 4;
            /// <summary> マスター：レコード開始行 </summary>
            public int recordStartRow = 5;
        }

        //----- field -----

        //----- property -----

        public MasterSettings Master { get; private set; }

        //----- method -----        

        public Settings()
        {
            var myAssembly = Assembly.GetEntryAssembly();

            var directory = Directory.GetParent(myAssembly.Location);

            var iniFilePath = Path.Combine(directory.FullName, "settings.ini");

            Master = IniFile.Read<MasterSettings>("Master", iniFilePath);
        }
    }
}
