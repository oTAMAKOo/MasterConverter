
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

        public static async Task ExportYamlRecords(string exportPath, string[] recordNames, object[] records)
        {
            var directory = PathUtility.Combine(Directory.GetParent(exportPath).FullName, Constants.RecordsFolderName);

            CreateCleanDirectory(directory);

            var tasks = new List<Task>();

            for (var i = 0; i < recordNames.Length; i++)
            {
                var fileName = recordNames[i].Trim();

                if (string.IsNullOrEmpty(fileName)) { continue; }

                var index = i;

                var task = Task.Run(() =>
                {
                    var filePath = PathUtility.Combine(directory, fileName + Constants.RecordFileExtension);

                    FileSystem.WriteFile(filePath, records[index], FileSystem.Format.Yaml);
                });

                tasks.Add(task);
            }

            await Task.WhenAll(tasks);
        }

        public static async Task ExportCellOption(string exportPath, RecordData[] records)
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

                    FileSystem.WriteFile(filePath, record.cells, FileSystem.Format.Yaml);
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
