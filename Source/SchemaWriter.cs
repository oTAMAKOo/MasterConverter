using System;
using System.IO;
using System.Threading;
using Extensions;
using MessagePack;
using YamlDotNet.Serialization;

namespace MasterConverter
{
    public class SchemaWriter
    {
        public static void ExportSchemaYaml(string exportPath, SerializeClass serializeClass)
        {
            var filePath = Path.ChangeExtension(exportPath, Constants.YamlMasterFileExtension);

            CreateFileDirectory(filePath);

            using (var file = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            {
                using (var writer = new StreamWriter(file))
                {
                    writer.Write(serializeClass.GetSchemaString());
                }
            }
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

            var serializer = new SerializerBuilder().Build();

            for (var i = 0; i < recordNames.Length; i++)
            {
                var fileName = recordNames[i].Trim();

                if (string.IsNullOrEmpty(fileName)) { continue; }
                
                var filePath = PathUtility.Combine(directory, fileName + Constants.RecordFileExtension);                

                using (var file = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                {
                    using (var writer = new StreamWriter(file))
                    {
                        serializer.Serialize(writer, records[i]);
                    }
                }
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