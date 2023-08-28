
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CommandLine;
using OfficeOpenXml;
using Extensions;

namespace MasterConverter
{
    class Program
    {
        private static Settings settings = null;

        private static bool autoExit = true;

        class CommandLineOptions
        {
            [Option("input", Required = true, HelpText = "Convert targets directorys.", Separator = ',', Default = new string[0])]
            public IEnumerable<string> Inputs { get; set; }
            [Option("mode", Required = true, HelpText = "Convert mode. (import or export).")]
            public string Mode { get; set; }
            [Option("exit", Required = false, HelpText = "Auto close on finish.")]
            public bool AutoExit { get; set; }
        }

        static void Main(string[] args)
        {
            /*=== 開発用 ========================================
        
            #if DEBUG

            var arguments = new List<string>();
            
            arguments.Add("--input");

            // 変換対象ディレクトリ ※ 「,」区切りで複数ディレクトリ指定可.
            arguments.Add(@"");
            
            arguments.Add("--mode");

            // 動作モード [import / export].
            arguments.Add("import");

            args = arguments.ToArray();

            #endif

            //==================================================*/

            // 引数.
            var options = Parser.Default.ParseArguments<CommandLineOptions>(args) as Parsed<CommandLineOptions>;

            if (options == null)
            {
                Exit("Arguments parse failed.");
            }

            // EPPlus License setup.
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // 設定ファイル読み込み.
            settings = new Settings();

            // 自動終了.
            autoExit = options.Value.AutoExit;

            // 対象取得.
            var inputs = options.Value.Inputs.ToArray();

            Console.WriteLine("\n--------- Target Directories ---------\n");

            inputs.ForEach(x => Console.WriteLine(" - {0}", x));

            Console.WriteLine("\n--------- Convert Processing ---------\n");

            TypeUtility.CreateTypeTable();

            var targets = inputs.SelectMany(x => FindClassSchemaDirectories(x))
                .Distinct()
                .OrderBy(x => x, new NaturalComparer())
                .ToArray();

            try
            {
                // メイン処理.
                var task = MainAsync(targets, options);

                task.Wait();

                Console.WriteLine("\nConvert Finished");
            }
            catch (Exception e)
            {
                Exit(e.ToString());
            }

            Exit();
        }

        static async Task MainAsync(string[] targets, Parsed<CommandLineOptions> options)
        {
            // フォーマット.

            var format = SerializationFileUtility.Format.Yaml;

            switch (settings.File.format)
            {
                case "yaml":
                    format = SerializationFileUtility.Format.Yaml;
                    break;

                case "json":
                    format = SerializationFileUtility.Format.Json;
                    break;
            }

            // メイン処理.

            var tasks = new List<Task>();

            foreach (var target in targets)
            {
                var directory = string.Empty;

                var pathType = PathUtility.GetFilePathType(target);

                switch (pathType)
                {
                    case PathUtility.FilePathType.Directory:
                        directory = target;
                        break;

                    case PathUtility.FilePathType.File:
                        directory = Path.GetDirectoryName(target);
                        break;
                }

                if (!Directory.Exists(directory))
                {
                    throw new DirectoryNotFoundException(string.Format("Directory not found. {0}", directory));
                }
                
                var task = Task.Run( async () =>
                {
                    try
                    {
                        var sw = System.Diagnostics.Stopwatch.StartNew();

                        switch (options.Value.Mode)
                        {
                            case "import":
                                await Import(directory, format);
                                break;

                            case "export":
                                await Export(directory, format);
                                break;

                            default:
                                throw new ArgumentException("Argument mode undefined.");
                        }

                        sw.Stop();

                        Console.WriteLine(" - {0} ({1:F2}ms)", target, sw.Elapsed.TotalMilliseconds);
                    }
                    catch
                    {
                        Console.WriteLine(" Failed: {0}", target);

                        throw;
                    }
                });

                tasks.Add(task);
            }

            if (tasks.Any())
            {
                await Task.WhenAll(tasks);
            }
        }

        private static void Exit(string message = null)
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

            Environment.Exit(string.IsNullOrEmpty(message) ? 0 : 1);
        }

        private static string[] FindClassSchemaDirectories(string directory)
        {
            var classSchemaExtension = Path.GetExtension(Constants.ClassSchemaFileName);

            return Directory.EnumerateFiles(directory, "*" + classSchemaExtension, SearchOption.AllDirectories)
                .Where(x => Path.GetFileName(x) == Constants.ClassSchemaFileName)
                .Select(x => Path.GetDirectoryName(x))
                .ToArray();
        }

        private static async Task Import(string directory, SerializationFileUtility.Format format)
        {
            var schemaFilePath = GetClassSchemaPath(directory);

            var recordDirectory = GetRecordFileDirectory(directory);

            var excelFilePath = GetEditExcelFilePath(directory);

            // 最終更新日確認.

            var isRequireImport = DataLoader.IsRequireImport(excelFilePath, recordDirectory);

            if (!isRequireImport){ return; }

            // クラス構成読み込み.
            var serializeClass = LoadClassSchema(directory, true);

            // インデックス情報読み込み.
            var indexData = DataLoader.LoadRecordIndex(excelFilePath, format);

            // レコード読み込み.            
            var records = await DataLoader.LoadRecords(recordDirectory, serializeClass.Class, format);

            // 編集用Excelを生成 + レコード入力.

            var fieldNameRow = settings.Master.fieldNameRow;
            var recordStartRow = settings.Master.recordStartRow;

            EditXlsxBuilder.Build(schemaFilePath, serializeClass, indexData, records, fieldNameRow, recordStartRow);
        }

        private static async Task Export(string directory, SerializationFileUtility.Format format)
        {
            var excelFilePath = GetEditExcelFilePath(directory);
            var recordDirectory = GetRecordFileDirectory(directory);

            // 最終更新日確認.

            var isRequireExport = DataLoader.IsRequireExport(excelFilePath, recordDirectory);

            if (!isRequireExport){ return; }

            // クラス構成読み込み.

            var serializeClass = LoadClassSchema(directory, true);

            // レコード取得.

            var fieldNameRow = settings.Master.fieldNameRow;
            var recordStartRow = settings.Master.recordStartRow;

            var records = DataLoader.LoadExcelRecords(excelFilePath, fieldNameRow, recordStartRow);

            // クラス変換.

            var instances = DeserializeRecords(serializeClass, records);
            
            // ディレクトリ作成.

            await DataWriter.CreateCleanDirectory(recordDirectory);

            // レコード出力.

            var recordNames = records.Select(x => x.recordName).ToArray();

            var tasks = new Task[] 
            {
                DataWriter.ExportRecords(excelFilePath, recordNames, instances, format),
                DataWriter.ExportCellOption(excelFilePath, records, format),
            };

            await Task.WhenAll(tasks);

            DataWriter.ExportRecordIndex(excelFilePath, recordNames, format);
        }

        // 編集用Excelファイルパス取得.
        private static string GetEditExcelFilePath(string directory)
        {
            var directoryInfo = new DirectoryInfo(directory);

            var filePath = PathUtility.Combine(directory, directoryInfo.Name + Constants.MasterFileExtension);

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
        private static SerializeClass LoadClassSchema(string directory, bool addIgnoreField)
        {
            var schemaFilePath = GetClassSchemaPath(directory);

            if (!File.Exists(schemaFilePath))
            {
                Exit(string.Format("File not found. {0}", schemaFilePath));
            }            

            var serializeClass = new SerializeClass();
            
            var dataTypeRow = settings.Master.dataTypeRow;
            var fieldNameRow = settings.Master.fieldNameRow;

            serializeClass.LoadClassSchema(schemaFilePath, dataTypeRow, fieldNameRow, addIgnoreField);

            return serializeClass;
        }

        // レコードをクラスにデシリアライズ. 
        private static object[] DeserializeRecords(SerializeClass serializeClass, RecordData[] records)
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