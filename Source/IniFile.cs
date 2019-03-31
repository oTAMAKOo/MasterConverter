
using System;
using System.IO;
using System.Text;
using System.Runtime.InteropServices;

public static class IniFile
{
    [DllImport("KERNEL32.DLL")]
    public static extern uint GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, uint nSize, string lpFileName);

    [DllImport("KERNEL32.DLL")]
    public static extern uint GetPrivateProfileInt(string lpAppName, string lpKeyName, int nDefault, string lpFileName);

    [DllImport("KERNEL32.DLL")]
    public static extern uint WritePrivateProfileString(string lpAppName, string lpKeyName, string lpString, string lpFileName);

    public static T Read<T>(string section, string filepath)
    {
        T ret = (T)Activator.CreateInstance(typeof(T));

        var sb = new StringBuilder(1024);

        foreach (var n in typeof(T).GetFields())
        {
            sb.Clear();

            if (n.FieldType == typeof(int))
            {
                n.SetValue(ret, (int)GetPrivateProfileInt(section, n.Name, 0, Path.GetFullPath(filepath)));
            }
            else if (n.FieldType == typeof(uint))
            {
                n.SetValue(ret, GetPrivateProfileInt(section, n.Name, 0, Path.GetFullPath(filepath)));
            }
            else if (n.FieldType == typeof(bool))
            {
                GetPrivateProfileString(section, n.Name, "false", sb, (uint)sb.Capacity, Path.GetFullPath(filepath));
                n.SetValue(ret, bool.Parse(sb.ToString()));
            }
            else
            {
                GetPrivateProfileString(section, n.Name, "", sb, (uint)sb.Capacity, Path.GetFullPath(filepath));
                n.SetValue(ret, sb.ToString());
            }
        };

        return ret;
    }

    public static void Write<T>(string secion, T data, string filepath)
    {
        foreach (var n in typeof(T).GetFields())
        {
            WritePrivateProfileString(secion, n.Name, n.GetValue(data).ToString(), Path.GetFullPath(filepath));
        };
    }
}