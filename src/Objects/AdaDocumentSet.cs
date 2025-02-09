using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XmlToExcel.Objects;

/// <summary>
/// Represents a set of Ada documents loaded from a source file.
/// </summary>
public class AdaDocumentSet : HashSet<AdaDocument>
{
    #region Members
    /// <summary>
    /// Gets or sets the path to the source file.
    /// </summary>
    public string SourceFilePath { get; set; }
    #endregion

    #region Constructors
    /// <summary>
    /// Initializes a new instance of the <see cref="AdaDocumentSet"/> class with the specified source file path.
    /// </summary>
    /// <param name="sourceFilePath">The path to the source file containing the Ada documents.</param>
    public AdaDocumentSet(string sourceFilePath)
    {
        SourceFilePath = sourceFilePath;
        if (File.Exists(sourceFilePath))
        {
            XDocument doc = XDocument.Load(sourceFilePath);
            foreach (var element in doc.Descendants("Ada"))
            {
                Add(AdaDocument.FromXmlElement(element));
            }
        }
    }
    #endregion

    #region Methods
    /// <summary>
    /// Gets a sorted set of Gimla types from the Ada documents.
    /// </summary>
    /// <returns>A sorted set of <see cref="GimlaType"/> objects.</returns>
    public SortedSet<GimlaType> GetGimlaTypes()
    {
        return new SortedSet<GimlaType>(this.Select(doc => new GimlaType
        {
            Code = doc.GimlaCode,
            Description = doc.GimlaDescription
        }));
    }
    #endregion
}
