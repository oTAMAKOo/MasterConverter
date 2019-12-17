﻿
using System;
using System.IO;
using System.Threading;
using Extensions;
using MessagePack;
using MessagePack.Resolvers;
using YamlDotNet.Serialization;

namespace MasterConverter
{
    public static class RecordWriter
    {
        //----- params -----

        //----- field -----

        //----- property -----

        //----- method -----

        public static void ExportMessagePack(string exportPath, object[] values, bool lz4Compress, string aesKey, string aesIv)
        {
            var filePath = Path.ChangeExtension(exportPath, Constants.MessagePackMasterFileExtension);
            
            var containerFormat = @"{{ ""records"": {0} }}";

            var valuesJson = JsonFx.Json.JsonWriter.Serialize(values);

            var messagePackJson = string.Format(containerFormat, valuesJson);
            
            byte[] bytes = null;

            if (lz4Compress)
            {
                var options = StandardResolverAllowPrivate.Options.WithCompression(MessagePackCompression.Lz4BlockArray);

                bytes = MessagePackSerializer.ConvertFromJson(messagePackJson, options);
            }
            else
            {
                bytes = MessagePackSerializer.ConvertFromJson(messagePackJson);
            }

            #if DEBUG

            Console.WriteLine("Json(LZ4MessagePack) :\n{0}\n", MessagePackSerializer.ConvertToJson(bytes));

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
            var filePath = Path.ChangeExtension(exportPath, Constants.YamlMasterFileExtension);

            var serializer = new SerializerBuilder().Build();

            CreateFileDirectory(filePath);

            using (var file = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            {
                using (var writer = new StreamWriter(file))
                {
                    serializer.Serialize(writer, records);
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
