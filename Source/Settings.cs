
using System.IO;
using System.Reflection;

namespace MasterConverter
{
    public sealed class Settings
    {
        //----- params -----
        
        public class MasterSettings
        {
            /// <summary> マスター：型名定義名行 </summary>
            public int dataTypeRow = 2;
            /// <summary> マスター：フィールド名定義行 </summary>
            public int fieldNameRow = 3;
            /// <summary> マスター：レコード開始行 </summary>
            public int recordStartRow = 4;
        }

        public class FileSettings
        {
            /// <summary> ファイルフォーマット </summary>
            public string format = "yaml";
        }

        //----- field -----

        //----- property -----

        public MasterSettings Master { get; private set; }

        public FileSettings File { get; private set; }

        //----- method -----        

        public Settings()
        {
            var myAssembly = Assembly.GetEntryAssembly();

            var directory = Directory.GetParent(myAssembly.Location);

            var iniFilePath = Path.Combine(directory.FullName, "settings.ini");

            Master = IniFile.Read<MasterSettings>("Rows", iniFilePath);

            File = IniFile.Read<FileSettings>("File", iniFilePath);
        }
    }
}
