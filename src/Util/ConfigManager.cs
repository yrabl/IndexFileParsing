using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XmlToExcel.Util;

/// <summary>
/// Provides methods to manage the configuration file for the Ada XML to Excel Converter application.
/// </summary>
public static class ConfigManager
{
    #region Members
    private static readonly string ConfigDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "AdaXmlToExcel");
    private static readonly string ConfigPath = Path.Combine(ConfigDirectory, "config.json");
    #endregion

    #region Methods
    /// <summary>
    /// Ensures the configuration directory and file exist. If they do not exist, they are created.
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
    /// Reads the content of the config.json file.
    /// </summary>
    /// <returns>A string containing the JSON content of the configuration file.</returns>
    public static string ReadConfig()
    {
        EnsureConfigExists();
        return File.ReadAllText(ConfigPath);
    }

    /// <summary>
    /// Writes the specified JSON data to the config.json file.
    /// </summary>
    /// <param name="json">The JSON data to write to the configuration file.</param>
    public static void WriteConfig(string json)
    {
        EnsureConfigExists();
        File.WriteAllText(ConfigPath, json);
    }
    #endregion
}
