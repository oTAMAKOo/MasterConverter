
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CommandLine;
using Extensions;
using MessagePack;

namespace MasterConverter
{
    class Program
    {
        private static Settings settings = null;
        private static bool autoExit = false;

        class CommandLineOptions
        {
            [Option("input", Required = true, HelpText = "Convert targets directorys.", Separator = ',', Default = new string[0])]
            public IEnumerable<string> Inputs { get; set; }
            [Option("mode", Required = true, HelpText = "Convert mode. (import or export or build).")]
            public string Mode { get; set; }
            [Option("tag", Required = false, HelpText = "Export target tags.", Separator = ',', Default = new string[0])]
            public IEnumerable<string> ExportTags { get; set; }
            [Option("export", Required = false, HelpText = "Export file. (messagepack or yaml or both).", Default = "both")]
            public string Export { get; set; }
            [Option("messagepack", Required = false, HelpText = "Messagepack export directory.", Default = null)]
            public string MessagePackDirectory { get; set; }
            [Option("yaml", Required = false, HelpText = "Yaml export directory.", Default = null)]
            public string YamlDirectory { get; set; }
            [Option("exit", Required = false, HelpText = "Auto close on finish.", Default = true)]
            public bool Exit { get; set; }
        }

        static void Main(string[] args)
        {
            var options = Parser.Default.ParseArguments<CommandLineOptions>(args) as Parsed<CommandLineOptions>;

            if (options == null)
            {
                Exit("Arguments parse failed.");
            }

            // 設定ファイル読み込み.
            settings = new Settings();

            // 自動終了.
            autoExit = options.Value.Exit;

            // 開発用.
            //options.Value.Inputs = new string[] { @"" };
            //options.Value.Mode = "import";

            foreach (var input in options.Value.Inputs)
            {
                var directory = string.Empty;

                var pathType = PathUtility.GetFilePathType(input);

                switch (pathType)
                {
                    case PathUtility.FilePathType.Directory:
                        directory = input;
                        break;

                    case PathUtility.FilePathType.File:
                        directory = Path.GetDirectoryName(input);
                        break;
                }

                try
                {
                    if (!Directory.Exists(directory))
                    {
                        throw new DirectoryNotFoundException(string.Format("Directory not found. {0}", directory));
                    }

                    switch (options.Value.Mode)
                    {
                        case "import":
                            Import(directory, options);
                            break;

                        case "export":
                            Export(directory, options);
                            break;

                        case "build":
                            Build(directory, options);
                            break;

                        default:
                            throw new ArgumentException("Argument mode undefined.");
                    }
                }
                catch (Exception e)
                {
                    Exit(e.ToString());
                }
            }

            #if DEBUG

            Console.ReadLine();

            #endif

            Environment.Exit(0);
        }

        private static void Exit(string message)
        {
            if (!string.IsNullOrEmpty(message))
            {
                Console.WriteLine(message);
            }

            #if DEBUG

            Console.ReadLine();

            #else

            if (!autoExit)
            {
                Console.ReadLine();
            }

            #endif

            Environment.Exit(1);
        }

        private static void Import(string directory, Parsed<CommandLineOptions> options)
        {
            var schemaFilePath = GetClassSchemaPath(directory);

            var fileDirectory = GetRecordFileDirectory(directory);

            // クラス構成読み込み.
            var serializeClass = LoadClassSchema(directory, null);

            // レコード読み込み.            
            var records = RecordLoader.LoadYamlRecords(fileDirectory, serializeClass.Class);

            // セルオプション読み込み.
            var cellOptions = CellOptionLoader.LoadYamlCellOptions(fileDirectory);

            // 編集用Excelを生成 + レコード入力.

            var fieldNameRow = settings.Master.fieldNameRow;
            var recordStartRow = settings.Master.recordStartRow;

            EditXlsxBuilder.Build(schemaFilePath, serializeClass, records, cellOptions, fieldNameRow, recordStartRow);
        }

        private static void Export(string directory, Parsed<CommandLineOptions> options)
        {
            var exportTags = options.Value.ExportTags.ToArray();

            // 出力ファイル名.

            var filePath = GetEditXlsxFilePath(directory);

            // 出力先ディレクトリ作成.

            RecordWriter.CreateCleanDirectory(filePath);
            
            // クラス構成読み込み.

            var serializeClass = LoadClassSchema(directory, exportTags);

            // レコード取得.

            var fieldNameRow = settings.Master.fieldNameRow;
            var recordStartRow = settings.Master.recordStartRow;

            var records = RecordLoader.LoadXlsxRecords(filePath, fieldNameRow, recordStartRow).OrderBy(x => x.index).ToArray();

            // クラス変換.

            var instances = DeserializeRecords(serializeClass, records);

            // レコード出力.

            var recordNames = records.Select(x => x.recordName).ToArray();
            
            RecordWriter.ExportYamlRecords(filePath, recordNames, instances);

            // セルオプション取得.

            var cellOptions = CellOptionLoader.LoadXlsxRecordsCellOptions(filePath, records);

            // セルオプション出力.

            CellOptionWriter.ExportYamlCellOptions(filePath, cellOptions);
        }

        private static void Build(string directory, Parsed<CommandLineOptions> options)
        {
            var exportTags = options.Value.ExportTags.ToArray();
            var export = options.Value.Export.ToLower();
            var messagepackDirectory = options.Value.MessagePackDirectory;
            var yamlDirectory = options.Value.YamlDirectory;

            if (string.IsNullOrEmpty(export))
            {
                throw new ArgumentException(string.Format("{0} is undefined.", nameof(export)));
            }

            var aesKey = settings.Export.AESKey;
            var aesIv = settings.Export.AESIv;
            var lz4compress = settings.Export.lz4compress;

            // 出力ファイル名.

            var filePath = GetEditXlsxFilePath(directory);

            // クラス構成読み込み.

            var serializeClass = LoadClassSchema(directory, exportTags);

            // レコード読み込み.

            var recordFileDirectory = GetRecordFileDirectory(directory);

            var records = RecordLoader.LoadYamlRecords(recordFileDirectory, serializeClass.Class);

            // クラス変換.

            var instances = DeserializeRecords(serializeClass, records);

            // MessagePack出力.

            if (export == "messagepack" || export == "both")
            {
                if (!string.IsNullOrEmpty(messagepackDirectory))
                {
                    filePath = PathUtility.Combine(messagepackDirectory, Path.GetFileName(filePath));
                }

                RecordWriter.ExportMessagePack(filePath, instances, lz4compress, aesKey, aesIv);
            }

            // Yaml出力.

            if (export == "yaml" || export == "both")
            {
                if (!string.IsNullOrEmpty(yamlDirectory))
                {
                    filePath = PathUtility.Combine(yamlDirectory, Path.GetFileName(filePath));
                }

                RecordWriter.ExportYaml(filePath, instances);
            }
        }

        // 編集用Excelファイルパス取得.
        private static string GetEditXlsxFilePath(string directory)
        {
            var directoryInfo = new DirectoryInfo(directory);

            var filePath = PathUtility.Combine(directory, string.Format("{0}.xlsx", directoryInfo.Name));

            return filePath;
        }

        // クラス情報ファイルのパス取得.
        private static string GetClassSchemaPath(string directory)
        {
            return PathUtility.Combine(directory, Constants.ClassSchemaFileName);
        }

        // レコードファイルのディレクトリ取得.
        private static string GetRecordFileDirectory(string directory)
        {
            return PathUtility.Combine(directory, Constants.RecordsFolderName);
        }

        // クラス構成読み込み.
        private static SerializeClass LoadClassSchema(string directory, string[] exportTags)
        {
            var schemaFilePath = GetClassSchemaPath(directory);

            if (!File.Exists(schemaFilePath))
            {
                Exit(string.Format("File not found. {0}", schemaFilePath));
            }            

            var serializeClass = new SerializeClass();

            var tagRow = settings.Master.tagRow;
            var dataTypeRow = settings.Master.dataTypeRow;
            var fieldNameRow = settings.Master.fieldNameRow;

            serializeClass.LoadClassSchema(schemaFilePath, exportTags, tagRow, dataTypeRow, fieldNameRow);

            return serializeClass;
        }

        // レコードをクラスにデシリアライズ.
        private static object[] DeserializeRecords(SerializeClass serializeClass, RecordLoader.RecordData[] records)
        {
            var list = new List<object>();

            foreach (var record in records)
            {
                var fieldValues = record.values.ToDictionary(x => x.fieldName, x => ConvertValueToText(x.value));

                var instance = serializeClass.CreateInstance(fieldValues);

                list.Add(instance);
            }

            return list.ToArray();
        }

        private static string ConvertValueToText(object value)
        {
            if (value == null) { return string.Empty; }

            var valueType = value.GetType();
            
            if (valueType.IsArray)
            {
                var enumerable = value as IEnumerable;

                if (enumerable == null) { return null; }

                var list = new List<string>();

                foreach (var element in enumerable)
                {
                    list.Add(element.ToString());
                }
                
                return string.Format("[{0}]", string.Join(",", list));
            }
            
            return value.ToString();
        }
    }
}