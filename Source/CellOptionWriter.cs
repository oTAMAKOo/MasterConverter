
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Extensions;
using YamlDotNet.Serialization;

namespace MasterConverter
{
    public static class CellOptionWriter
    {
        //----- params -----

        //----- field -----

        //----- property -----

        //----- method -----

        public static void ExportYamlCellOptions(string exportPath, CellOptionLoader.CellOption[] options)
        {
            var directory = PathUtility.Combine(Directory.GetParent(exportPath).FullName, Constants.RecordsFolderName);

            var serializer = new SerializerBuilder().Build();

            foreach (var option in options)
            {
                if (option == null) { continue; }

                if (option.cellInfos.All(x => x == null)) { continue; }

                var fileName = option.recordName.Trim();

                if (string.IsNullOrEmpty(fileName)) { continue; }

                var filePath = PathUtility.Combine(directory, fileName + Constants.CellOptionFileExtension);

                using (var file = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                {
                    using (var writer = new StreamWriter(file))
                    {
                        serializer.Serialize(writer, option);
                    }
                }
            }
        }
    }
}
