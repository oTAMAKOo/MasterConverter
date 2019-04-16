
using System;
using System.IO;
using IniParser.Model;

public static class IniFile
{
    public static T Read<T>(string section, string filepath)
    {
        T ret = (T)Activator.CreateInstance(typeof(T));

        var parser = new IniParser.Parser.IniDataParser();
        IniData config = parser.Parse(File.ReadAllText(filepath));
        KeyDataCollection sectionData =  config.Sections[section];
        
        foreach (var n in typeof(T).GetFields())
        {
            if (n.FieldType == typeof(int))
            {
                if (sectionData.ContainsKey(n.Name))
                {
                    var value = sectionData[n.Name];
                    n.SetValue(ret, int.Parse(value));
                }
            }
            else if (n.FieldType == typeof(bool))
            {
                if (sectionData.ContainsKey(n.Name))
                {
                    var value = sectionData[n.Name];
                    n.SetValue(ret, bool.Parse(value));
                }
            }
            else if (n.FieldType == typeof(string))
            {
                if (sectionData.ContainsKey(n.Name))
                {
                    var value = sectionData[n.Name];
                    n.SetValue(ret, value);
                }
            }
            
        };

        return ret;
    }
}