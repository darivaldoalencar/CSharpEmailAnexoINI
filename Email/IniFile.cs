using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace Email
{
    public class IniFile
    {
        public string Path { get; set; }

        public IniFile()
        {
            this.Path = "C:/ProjetosWeb/CSharpEmailAnexoINI/Email/configuracoes/CONFIGURACOES.INI"; 
        }

        public void WriteValue(string section, string key, string value)
        {
            NativeMethods.WritePrivateProfileString(section, key, value, Path);
        }

        public string ReadValue(string section, string key, string Default)
        {
            var buffer = new StringBuilder(255);
            NativeMethods.GetPrivateProfileString(section, key, Default, buffer, 255, Path);

            return buffer.ToString();
        }

        public void WriteValue(string section, string key, int value)
        {
            NativeMethods.WritePrivateProfileString(section, key, value.ToString(CultureInfo.InvariantCulture), Path);
        }

        public int ReadValue(string section, string key, int Default)
        {
            var buffer = new StringBuilder(255);
            NativeMethods.GetPrivateProfileString(section, key, Default.ToString(CultureInfo.InvariantCulture), buffer, 255, Path);

            return int.Parse(buffer.ToString());
        }

        public void WriteValue(string section, string key, UInt16 value)
        {
            NativeMethods.WritePrivateProfileString(section, key, value.ToString(CultureInfo.InvariantCulture), Path);
        }
        public UInt16 ReadValue(string section, string key, UInt16 Default)
        {
            var buffer = new StringBuilder(255);
            NativeMethods.GetPrivateProfileString(section, key, Default.ToString(CultureInfo.InvariantCulture), buffer, 255, Path);

            return UInt16.Parse(buffer.ToString());
        }
        public void WriteValue(string section, string key, UInt32 value)
        {
            NativeMethods.WritePrivateProfileString(section, key, value.ToString(CultureInfo.InvariantCulture), Path);
        }
        public UInt32 ReadValue(string section, string key, UInt32 Default)
        {
            var buffer = new StringBuilder(255);
            NativeMethods.GetPrivateProfileString(section, key, Default.ToString(CultureInfo.InvariantCulture), buffer, 255, Path);

            return UInt32.Parse(buffer.ToString());
        }

        public bool ReadValue(string section, string key, bool Default)
        {
            var buffer = new StringBuilder(255);
            NativeMethods.GetPrivateProfileString(section, key, Default.ToString(CultureInfo.InvariantCulture), buffer, 255, Path);

            return (buffer.ToString() != "False");
        }
        public void WriteValue(string section, string key, bool value)
        {
            NativeMethods.WritePrivateProfileString(section, key, value.ToString(CultureInfo.InvariantCulture), Path);
        }

        static class NativeMethods
        {
            [DllImport("kernel32", CharSet = CharSet.Unicode)]
            public static extern uint WritePrivateProfileString(string section,
            string key, string val, string filePath);

            [DllImport("kernel32", CharSet = CharSet.Unicode)]
            public static extern int GetPrivateProfileString(string section,
            string key, string def, StringBuilder retVal, int size,
            string filePath);
        }

        public string getTAG(string tag)
        {
            var leitura = new IniFile(this.Path).ReadValue("CONFIGURACOES", tag, string.Empty); ;                 
            return leitura;
        }

    }
}
