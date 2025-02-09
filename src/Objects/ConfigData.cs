namespace XmlToExcel.Objects;

/// <summary>
/// Represents the configuration data for the XML to Excel conversion process.
/// </summary>
public class ConfigData
{
    #region Members
    /// <summary>
    /// Gets or sets the path to the data files.
    /// </summary>
    public string DataPath { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the path to the Excel file.
    /// </summary>
    public string ExcelFile { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets a value indicating whether to delete files after processing.
    /// </summary>
    public bool DeleteFiles { get; set; } = false;

    /// <summary>
    /// Gets or sets a value indicating whether to rename files after processing.
    /// </summary>
    public bool RenameFiles { get; set; } = false;
    #endregion

    #region Constructors
    /// <summary>
    /// Initializes a new instance of the <see cref="ConfigData"/> class.
    /// </summary>
    public ConfigData()
    {

    }
    #endregion
}
