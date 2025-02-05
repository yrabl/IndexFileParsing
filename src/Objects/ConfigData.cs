namespace XmlToExcel.Objects;

public class ConfigData
{
    #region Members
    public string DataPath { get; set; } = string.Empty;
    public string ExcelFile { get; set; } = string.Empty;
    public bool DeleteFiles { get; set; } = false;
    #endregion

    #region Constructors
    public ConfigData()
    {

    }
    #endregion
}
