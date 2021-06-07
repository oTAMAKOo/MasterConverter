
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Extensions;

namespace MasterConverter
{
    public static class DataWriter
    {
        //----- params -----

        //----- field -----
        
        //----- property -----

        //----- method -----

        public static void ExportRecordIndex(string excelFilePath, string[] recordNames, SerializationFileUtility.Format format)
        {
            var filePath = Path.ChangeExtension(excelFilePath, Constants.IndexFileExtension);

            var indexData = new IndexData()
            {
                records = recordNames
            };

            SerializationFileUtility.WriteFile(filePath, indexData, format);
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

        public static async Task CreateCleanDirectory(string exportPath)
        {
            var directory = PathUtility.Combine(Directory.GetParent(exportPath).FullName, Constants.RecordsFolderName);
           
            if (Directory.Exists(directory))
            {
                DirectoryUtility.Delete(directory);

                // ディレクトリの削除は非同期で実行される為、削除完了するまで待機する.
                while (Directory.Exists(directory))
                {
                    await Task.Delay(TimeSpan.FromMilliseconds(50));
                }

                await Task.Delay(TimeSpan.FromMilliseconds(50));
            }

            Directory.CreateDirectory(directory);            
        }

        public static async Task ExportRecords(string exportPath, string[] recordNames, object[] records, SerializationFileUtility.Format format)
        {
            var directory = PathUtility.Combine(Directory.GetParent(exportPath).FullName, Constants.RecordsFolderName);

            await CreateCleanDirectory(directory);

            var tasks = new List<Task>();

            for (var i = 0; i < recordNames.Length; i++)
            {
                var fileName = recordNames[i].Trim();

                if (string.IsNullOrEmpty(fileName)) { continue; }

                var index = i;

                var task = Task.Run(() =>
                {
                    var filePath = PathUtility.Combine(directory, fileName + Constants.RecordFileExtension);

                    SerializationFileUtility.WriteFile(filePath, records[index], format);
                });

                tasks.Add(task);
            }

            await Task.WhenAll(tasks);
        }

        public static async Task ExportCellOption(string exportPath, RecordData[] records, SerializationFileUtility.Format format)
        {
            var directory = PathUtility.Combine(Directory.GetParent(exportPath).FullName, Constants.RecordsFolderName);

            var tasks = new List<Task>();

            foreach (var record in records)
            {
                if (record.cells == null){ continue; }

                if (record.cells.IsEmpty()){ continue; }

                var task = Task.Run(() =>
                {
                    var filePath = PathUtility.Combine(directory, record.recordName + Constants.CellOptionFileExtension);

                    SerializationFileUtility.WriteFile(filePath, record.cells, format);
                });

                tasks.Add(task);
            }

            await Task.WhenAll(tasks);
        }

        private static void CreateFileDirectory(string filePath)
        {
            var directory = Path.GetDirectoryName(filePath);

            if (Directory.Exists(directory)) { return; }
            
            Directory.CreateDirectory(directory);            
        }
    }
}
