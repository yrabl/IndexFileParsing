using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace XmlToExcel.Objects;

/// <summary>
/// Represents an Ada document with various properties and methods for handling document data.
/// </summary>
public class AdaDocument : IEquatable<AdaDocument?>
{
    #region Members
    /// <summary>
    /// Gets or sets the document ID.
    /// </summary>
    [DisplayName("Document ID")]
    public long DocumentAdaId { get; set; }

    /// <summary>
    /// Gets or sets the document date.
    /// </summary>
    [DisplayName("Document Date")]
    public DateOnly DocumentDate { get; set; }

    /// <summary>
    /// Gets or sets the Gimla code.
    /// </summary>
    [DisplayName("Gimla Code")]
    public int GimlaCode { get; set; }

    /// <summary>
    /// Gets or sets the Gimla description.
    /// </summary>
    [DisplayName("Gimla Description")]
    public string GimlaDescription { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the document type.
    /// </summary>
    [DisplayName("Type")]
    public int DocumentType { get; set; }

    /// <summary>
    /// Gets or sets the document type description.
    /// </summary>
    [DisplayName("Description")]
    public string DocumentTypeDescription { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the event date.
    /// </summary>
    [DisplayName("Event Date")]
    public DateOnly? EventDate { get; set; }

    /// <summary>
    /// Gets the new file name based on the document properties.
    /// </summary>
    [DisplayName("New File Name")]
    public string NewFileName
    {
        get
        {
            string sanitizedDocTypeDesc = GetSanitizedDocTypeDesc(DocumentTypeDescription);
            return $"{sanitizedDocTypeDesc}-{DocumentType}-{DocumentDate.ToString("yyyy-MM-dd")}-{DocumentAdaId}";
        }
    }
    #endregion

    #region Methods
    /// <summary>
    /// Creates an <see cref="AdaDocument"/> instance from an XML element.
    /// </summary>
    /// <param name="element">The XML element containing the document data.</param>
    /// <returns>An <see cref="AdaDocument"/> instance.</returns>
    public static AdaDocument FromXmlElement(XElement element)
    {
        string dateFormat = "yyyy-MM-ddTHH:mm:ss";
        AdaDocument adaDocument = new AdaDocument
        {
            DocumentAdaId = long.Parse(element.Element("doc_ada_id").Value),
            DocumentDate = DateOnly.ParseExact(element.Element("doc_date").Value, dateFormat),
            GimlaCode = int.Parse(element.Element("gimla_code").Value),
            GimlaDescription = element.Element("gimal_desc").Value ?? string.Empty,
            DocumentType = int.Parse(element.Element("doc_type").Value),
            DocumentTypeDescription = element.Element("doc_type_desc").Value ?? string.Empty,
            EventDate = DateOnly.TryParseExact(element.Element("event_date")?.Value, dateFormat, out var eventDate) ? eventDate : null
        };
        return adaDocument;
    }

    /// <summary>
    /// Determines whether the specified object is equal to the current object.
    /// </summary>
    /// <param name="obj">The object to compare with the current object.</param>
    /// <returns>true if the specified object is equal to the current object; otherwise, false.</returns>
    public override bool Equals(object? obj)
    {
        return Equals(obj as AdaDocument);
    }

    /// <summary>
    /// Determines whether the specified <see cref="AdaDocument"/> is equal to the current <see cref="AdaDocument"/>.
    /// </summary>
    /// <param name="other">The <see cref="AdaDocument"/> to compare with the current <see cref="AdaDocument"/>.</param>
    /// <returns>true if the specified <see cref="AdaDocument"/> is equal to the current <see cref="AdaDocument"/>; otherwise, false.</returns>
    public bool Equals(AdaDocument? other)
    {
        return other is not null &&
               DocumentAdaId == other.DocumentAdaId &&
               DocumentDate.Equals(other.DocumentDate) &&
               GimlaCode == other.GimlaCode &&
               GimlaDescription == other.GimlaDescription &&
               DocumentType == other.DocumentType &&
               DocumentTypeDescription == other.DocumentTypeDescription &&
               EqualityComparer<DateOnly?>.Default.Equals(EventDate, other.EventDate);
    }

    /// <summary>
    /// Serves as the default hash function.
    /// </summary>
    /// <returns>A hash code for the current object.</returns>
    public override int GetHashCode()
    {
        return HashCode.Combine(DocumentAdaId, DocumentDate, GimlaCode, GimlaDescription, DocumentType, DocumentTypeDescription, EventDate);
    }

    /// <summary>
    /// Determines whether two specified instances of <see cref="AdaDocument"/> are equal.
    /// </summary>
    /// <param name="left">The first <see cref="AdaDocument"/> to compare.</param>
    /// <param name="right">The second <see cref="AdaDocument"/> to compare.</param>
    /// <returns>true if the two <see cref="AdaDocument"/> instances are equal; otherwise, false.</returns>
    public static bool operator ==(AdaDocument? left, AdaDocument? right)
    {
        return EqualityComparer<AdaDocument>.Default.Equals(left, right);
    }

    /// <summary>
    /// Determines whether two specified instances of <see cref="AdaDocument"/> are not equal.
    /// </summary>
    /// <param name="left">The first <see cref="AdaDocument"/> to compare.</param>
    /// <param name="right">The second <see cref="AdaDocument"/> to compare.</param>
    /// <returns>true if the two <see cref="AdaDocument"/> instances are not equal; otherwise, false.</returns>
    public static bool operator !=(AdaDocument? left, AdaDocument? right)
    {
        return !(left == right);
    }

    /// <summary>
    /// Gets the matching file in the specified path based on the document ID.
    /// </summary>
    /// <param name="path">The path to search for the file.</param>
    /// <returns>The matching file path, or null if no matching file is found.</returns>
    public string GetMatchingFileInPath(string path)
    {
        return Directory.GetFiles(path, $"{DocumentAdaId}.*").FirstOrDefault();
    }

    /// <summary>
    /// Gets the renamed file in the specified path based on the new file name.
    /// </summary>
    /// <param name="path">The path to search for the file.</param>
    /// <returns>The renamed file path, or null if no renamed file is found.</returns>
    public string GetRenamedFileInPath(string path)
    {
        return Directory.GetFiles(path, $"{NewFileName}.*").FirstOrDefault();
    }

    /// <summary>
    /// Sanitizes the document type description by replacing spaces with underscores and removing extra characters.
    /// </summary>
    /// <param name="docTypeDesc">The document type description to sanitize.</param>
    /// <returns>The sanitized document type description.</returns>
    private static string GetSanitizedDocTypeDesc(string docTypeDesc)
    {
        string sanitizedDocTypeDesc = System.Text.RegularExpressions.Regex.Replace(docTypeDesc.Trim(), @"[ \\/]+", "_");
        sanitizedDocTypeDesc = System.Text.RegularExpressions.Regex.Replace(sanitizedDocTypeDesc, @"_-_+", "-");
        return sanitizedDocTypeDesc;
    }
    #endregion
}
