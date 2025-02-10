using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XmlToExcel.Util;

public static class ConfigManager
{
    #region Members
    private static readonly string ConfigDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "AdaXmlToExcel");
    private static readonly string ConfigPath = Path.Combine(ConfigDirectory, "config.json");
    #endregion

    #region Methods
    /// <summary>
    /// Ensures the configuration directory and file exist.
    /// </summary>
    public static void EnsureConfigExists()
    {
        if (!Directory.Exists(ConfigDirectory))
        {
            Directory.CreateDirectory(ConfigDirectory);
        }

        if (!File.Exists(ConfigPath))
        {
            File.WriteAllText(ConfigPath, "{\"DataPath\":\"\",\"ExcelFile\":\"\",\"DeleteFiles\":false,\"RenameFiles\":false}"); // Initialize with empty JSON
        }
    }

    /// <summary>
    /// Reads the config.json file.
    /// </summary>
    public static string ReadConfig()
    {
        EnsureConfigExists();
        return File.ReadAllText(ConfigPath);
    }

    /// <summary>
    /// Writes data to config.json.
    /// </summary>
    public static void WriteConfig(string json)
    {
        EnsureConfigExists();
        File.WriteAllText(ConfigPath, json);
    }
    #endregion
}
