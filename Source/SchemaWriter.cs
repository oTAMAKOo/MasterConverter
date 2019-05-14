using System.IO;

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

        private static void CreateFileDirectory(string filePath)
        {
            var directory = Path.GetDirectoryName(filePath);

            if (Directory.Exists(directory)) { return; }
            
            Directory.CreateDirectory(directory);            
        }        
    }
}