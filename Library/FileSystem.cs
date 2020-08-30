
using System.IO;
using System.Text;
using Newtonsoft.Json;
using YamlDotNet.Serialization;

namespace Extensions
{
    public static class FileSystem
    {
        public enum Format
        {
            Yaml,
            Json,
        }

        public static void WriteFile<T>(string filePath, T value, Format format)
        {
            var directory = Path.GetDirectoryName(filePath);

            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using (var file = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            {
                using (var writer = new StreamWriter(file, new UTF8Encoding(false)))
                {
                    switch (format)
                    {
                        case Format.Json:
                            {
                                using (var jsonTextWriter = new JsonTextWriter(writer))
                                {
                                    jsonTextWriter.Formatting = Formatting.Indented;

                                    var jsonSerializer = new JsonSerializer()
                                    {
                                        Formatting = Formatting.Indented,
                                        NullValueHandling = NullValueHandling.Ignore,
                                    };

                                    jsonSerializer.Serialize(jsonTextWriter, value);
                                }
                            }
                            break;

                        case Format.Yaml:
                            {
                                var builder = new SerializerBuilder();

                                builder.ConfigureDefaultValuesHandling(DefaultValuesHandling.OmitNull);

                                var yamlSerializer = builder.Build();

                                yamlSerializer.Serialize(writer, value);
                            }
                            break;
                    }
                }
            }
        }

        public static T LoadFile<T>(string filePath, Format format) where T : class 
        {
            if (!File.Exists(filePath)) { return null; }

            T result = null;

            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var reader = new StreamReader(file, new UTF8Encoding(false)))
                {
                    switch (format)
                    {
                        case Format.Json:
                            {
                                using (var jsonTextReader = new JsonTextReader(reader))
                                {
                                    var jsonSerializer = new JsonSerializer()
                                    {
                                        Formatting = Formatting.Indented,
                                        NullValueHandling = NullValueHandling.Ignore,
                                    };

                                    result = jsonSerializer.Deserialize<T>(jsonTextReader);
                                }
                            }
                            break;

                        case Format.Yaml:
                            {
                                var contents = reader.ReadToEnd();

                                var builder = new DeserializerBuilder();

                                builder.IgnoreUnmatchedProperties();

                                var yamlDeserializer = builder.Build();
                                
                                result = yamlDeserializer.Deserialize<T>(contents);
                            }
                            break;
                    }
                }
            }

            return result;
        }
    }
}
