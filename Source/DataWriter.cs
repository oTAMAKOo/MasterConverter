
using System;
using System.IO;
using System.Threading;
using Extensions;
using MessagePack;
using MessagePack.Resolvers;

namespace MasterConverter
{
    public static class DataWriter
    {
        //----- params -----

        //----- field -----

        //----- property -----

        //----- method -----

        public static void ExportRecordIndex(string excelFilePath, string[] recordNames)
        {
            var filePath = Path.ChangeExtension(excelFilePath, Constants.IndexFileExtension);

            var indexData = new IndexData()
            {
                records = recordNames
            };

            FileSystem.WriteFile(filePath, indexData, FileSystem.Format.Yaml);
        }

        private static string GetExportPath(string filePath, string extension)
        {
            const string MasterSuffix = "Master";

            var exportPath = string.Empty;

            var directory = Path.GetDirectoryName(filePath);
            var fileName = Path.GetFileNameWithoutExtension(filePath);

            if (fileName.EndsWith(MasterSuffix))
            {
                fileName = fileName.SafeSubstring(0, fileName.Length - MasterSuffix.Length);
            }

            exportPath = PathUtility.Combine(directory, fileName) + extension;

            return exportPath;
        }

        public static void ExportMessagePack(string exportPath, Type dataType, object[] values, bool lz4Compress, string aesKey, string aesIv)
        {
            var filePath = GetExportPath(exportPath, Constants.MessagePackMasterFileExtension);

            //----- データ格納用のコンテナクラス作成 -----

            var containerClassType = typeof(ContainerClass<>);

            var containerType = containerClassType.MakeGenericType(dataType);

            var containerInstance = Activator.CreateInstance(containerType);

            var container = containerInstance as IContainerClass;

            container.SetRecords(values);

            //----- MessagePackのバイナリデータ出力 -----

            var options = StandardResolverAllowPrivate.Options.WithResolver(MessagePackContractResolver.Instance);

            if (lz4Compress)
            {
                options = options.WithCompression(MessagePackCompression.Lz4BlockArray);
            }

            var bytes = MessagePackSerializer.Serialize(containerInstance, options);

            #if DEBUG

            Console.WriteLine("Json :\n{0}\n", MessagePackSerializer.ConvertToJson(bytes));

            #endif


            if (!string.IsNullOrEmpty(aesKey) && !string.IsNullOrEmpty(aesIv))
            {
                var aesManaged = AESExtension.CreateAesManaged(aesKey, aesIv);

                bytes = bytes.Encrypt(aesManaged);
            }

            CreateFileDirectory(filePath);

            using (var file = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            {
                file.Write(bytes, 0, bytes.Length);
            }
        }

        public static void ExportYaml(string exportPath, object[] records)
        {
            var filePath = GetExportPath(exportPath, Constants.YamlMasterFileExtension);
            
            CreateFileDirectory(filePath);

            FileSystem.WriteFile(filePath, records, FileSystem.Format.Yaml);
        }

        public static void CreateCleanDirectory(string exportPath)
        {
            var directory = PathUtility.Combine(Directory.GetParent(exportPath).FullName, Constants.RecordsFolderName);
           
            if (Directory.Exists(directory))
            {
                DirectoryUtility.Delete(directory);

                // ディレクトリの削除は非同期で実行される為、削除完了するまで待機する.
                while (Directory.Exists(directory))
                {
                    Thread.Sleep(100);
                }
            }

            Directory.CreateDirectory(directory);            
        }

        public static void ExportYamlRecords(string exportPath, string[] recordNames, object[] records)
        {
            var directory = PathUtility.Combine(Directory.GetParent(exportPath).FullName, Constants.RecordsFolderName);

            CreateCleanDirectory(directory);

            for (var i = 0; i < recordNames.Length; i++)
            {
                var fileName = recordNames[i].Trim();

                if (string.IsNullOrEmpty(fileName)) { continue; }
                
                var filePath = PathUtility.Combine(directory, fileName + Constants.RecordFileExtension);                
                
                FileSystem.WriteFile(filePath, records[i], FileSystem.Format.Yaml);
            }
        }

        public static void ExportCellOption(string exportPath, RecordData[] records)
        {
            var directory = PathUtility.Combine(Directory.GetParent(exportPath).FullName, Constants.RecordsFolderName);

            foreach (var record in records)
            {
                if (record.cells == null){ continue; }

                if (record.cells.IsEmpty()){ continue; }

                var filePath = PathUtility.Combine(directory, record.recordName + Constants.CellOptionFileExtension);

                FileSystem.WriteFile(filePath, record.cells, FileSystem.Format.Yaml);
            }
        }

        private static void CreateFileDirectory(string filePath)
        {
            var directory = Path.GetDirectoryName(filePath);

            if (Directory.Exists(directory)) { return; }
            
            Directory.CreateDirectory(directory);            
        }
    }
}
