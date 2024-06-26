﻿
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
            
            var directory = Path.GetDirectoryName(filePath);
            var fileName = Path.GetFileNameWithoutExtension(filePath);

            if (fileName.EndsWith(MasterSuffix))
            {
                fileName = fileName.SafeSubstring(0, fileName.Length - MasterSuffix.Length);
            }

            var exportPath = PathUtility.Combine(directory, fileName) + extension;

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

            var tasks = new List<Task>();

            for (var i = 0; i < recordNames.Length; i++)
            {
                var index = i;
                var recordName = recordNames[index];
                var record = records[index];

                var task = Task.Run(async () =>
                {
                    var fileName = recordName.Trim();

                    if (string.IsNullOrEmpty(fileName)) { return; }

                    var filePath = PathUtility.Combine(directory, fileName + Constants.RecordFileExtension);

                    while (FileUtility.IsFileLocked(filePath))
                    {
                        await Task.Delay(1);
                    }

                    SerializationFileUtility.WriteFile(filePath, record, format);
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
                var cells = record.cells;
                var recordName = record.recordName;

                if (cells == null){ continue; }

                if (cells.IsEmpty()){ continue; }

                var task = Task.Run(async () =>
                {
                    var filePath = PathUtility.Combine(directory, recordName + Constants.CellOptionFileExtension);

                    while (FileUtility.IsFileLocked(filePath))
                    {
                        await Task.Delay(1);
                    }

                    SerializationFileUtility.WriteFile(filePath, cells, format);
                });

                tasks.Add(task);
            }

            await Task.WhenAll(tasks);
        }
    }
}
