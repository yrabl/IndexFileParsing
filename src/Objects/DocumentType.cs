namespace XmlToExcel.Objects;

/// <summary>
/// Represents a type of document with a code and description.
/// </summary>
public class DocumentType : IEquatable<DocumentType?>, IComparable<DocumentType>
{
    #region Members
    /// <summary>
    /// Gets or sets the code of the document type.
    /// </summary>
    public int Code { get; set; }

    /// <summary>
    /// Gets or sets the description of the document type.
    /// </summary>
    public string Description { get; set; } = null!;
    #endregion

    #region Methods
    /// <summary>
    /// Determines whether the specified object is equal to the current object.
    /// </summary>
    /// <param name="obj">The object to compare with the current object.</param>
    /// <returns>true if the specified object is equal to the current object; otherwise, false.</returns>
    public override bool Equals(object? obj)
    {
        return Equals(obj as DocumentType);
    }

    /// <summary>
    /// Determines whether the specified DocumentType is equal to the current DocumentType.
    /// </summary>
    /// <param name="other">The DocumentType to compare with the current DocumentType.</param>
    /// <returns>true if the specified DocumentType is equal to the current DocumentType; otherwise, false.</returns>
    public bool Equals(DocumentType? other)
    {
        return other is not null &&
               Code == other.Code;
    }

    /// <summary>
    /// Serves as the default hash function.
    /// </summary>
    /// <returns>A hash code for the current object.</returns>
    public override int GetHashCode()
    {
        return HashCode.Combine(Code);
    }

    /// <summary>
    /// Determines whether two specified instances of DocumentType are equal.
    /// </summary>
    /// <param name="left">The first DocumentType to compare.</param>
    /// <param name="right">The second DocumentType to compare.</param>
    /// <returns>true if the two DocumentType instances are equal; otherwise, false.</returns>
    public static bool operator ==(DocumentType? left, DocumentType? right)
    {
        return EqualityComparer<DocumentType>.Default.Equals(left, right);
    }

    /// <summary>
    /// Determines whether two specified instances of DocumentType are not equal.
    /// </summary>
    /// <param name="left">The first DocumentType to compare.</param>
    /// <param name="right">The second DocumentType to compare.</param>
    /// <returns>true if the two DocumentType instances are not equal; otherwise, false.</returns>
    public static bool operator !=(DocumentType? left, DocumentType? right)
    {
        return !(left == right);
    }

    /// <summary>
    /// Compares the current instance with another object of the same type.
    /// </summary>
    /// <param name="other">An object to compare with this instance.</param>
    /// <returns>A value that indicates the relative order of the objects being compared.</returns>
    public int CompareTo(DocumentType? other)
    {
        if (other == null) return 1;
        return Code.CompareTo(other.Code);
    }
    #endregion
}
