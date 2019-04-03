using System;

namespace MasterConverter
{
    public static class Constants
    {
        /// <summary> クラス情報定義ファイル名 </summary>
        public const string ClassSchemaFileName = "ClassSchema.xlsx";

        /// <summary> 編集用マスターファイル拡張子  </summary>
        public const string MasterFileExtension = ".xlsx";

        /// <summary> マスターデータシート名 </summary>
        public const string MasterSheetName = "Master";

        /// <summary> マスターファイル(Yaml)拡張子 </summary>
        public const string YamlMasterFileExtension = ".yml";

        /// <summary> マスターファイル(MessagePack)拡張子 </summary>
        public const string MessagePackMasterFileExtension = ".master";

        /// <summary> レコードフォルダ名 </summary>
        public const string RecordsFolderName = "Records";

        /// <summary> レコードファイル拡張子 </summary>
        public const string RecordFileExtension = ".yml"; // TODO: .recordに変える
    }
}
