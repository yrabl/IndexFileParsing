
namespace XmlToExcel.Objects;

/// <summary>
/// Represents a mapping between a Gimla code and a document type.
/// </summary>
public class GimlaToDocument : IEquatable<GimlaToDocument?>, IComparable<GimlaToDocument>
{
    #region Members
    /// <summary>
    /// Gets or sets the Gimla code.
    /// </summary>
    [DisplayName("Gimla Code")]
    public int GimlaCode { get; set; }

    /// <summary>
    /// Gets or sets the description of the Gimla code.
    /// </summary>
    [DisplayName("Gimla Description")]
    public string GimlaDescription { get; set; } = null!;

    /// <summary>
    /// Gets or sets the document type.
    /// </summary>
    [DisplayName("Doc Type")]
    public int DocType { get; set; }

    /// <summary>
    /// Gets or sets the description of the document type.
    /// </summary>
    [DisplayName("Doc Description")]
    public string DocDescription { get; set; } = null!;
    #endregion

    #region Methods
    /// <summary>
    /// Determines whether the specified object is equal to the current object.
    /// </summary>
    /// <param name="obj">The object to compare with the current object.</param>
    /// <returns>true if the specified object is equal to the current object; otherwise, false.</returns>
    public override bool Equals(object? obj)
    {
        return obj is GimlaToDocument other && Equals(other);
    }

    /// <summary>
    /// Determines whether the specified <see cref="GimlaToDocument"/> is equal to the current <see cref="GimlaToDocument"/>.
    /// </summary>
    /// <param name="other">The <see cref="GimlaToDocument"/> to compare with the current <see cref="GimlaToDocument"/>.</param>
    /// <returns>true if the specified <see cref="GimlaToDocument"/> is equal to the current <see cref="GimlaToDocument"/>; otherwise, false.</returns>
    public bool Equals(GimlaToDocument? other)
    {
        return other is not null &&
               GimlaCode == other.GimlaCode &&
               DocType == other.DocType;
    }

    /// <summary>
    /// Serves as the default hash function.
    /// </summary>
    /// <returns>A hash code for the current object.</returns>
    public override int GetHashCode()
    {
        return HashCode.Combine(GimlaCode, DocType);
    }

    /// <summary>
    /// Determines whether two specified instances of <see cref="GimlaToDocument"/> are equal.
    /// </summary>
    /// <param name="left">The first <see cref="GimlaToDocument"/> to compare.</param>
    /// <param name="right">The second <see cref="GimlaToDocument"/> to compare.</param>
    /// <returns>true if the two <see cref="GimlaToDocument"/> instances are equal; otherwise, false.</returns>
    public static bool operator ==(GimlaToDocument? left, GimlaToDocument? right)
    {
        return EqualityComparer<GimlaToDocument>.Default.Equals(left, right);
    }

    /// <summary>
    /// Determines whether two specified instances of <see cref="GimlaToDocument"/> are not equal.
    /// </summary>
    /// <param name="left">The first <see cref="GimlaToDocument"/> to compare.</param>
    /// <param name="right">The second <see cref="GimlaToDocument"/> to compare.</param>
    /// <returns>true if the two <see cref="GimlaToDocument"/> instances are not equal; otherwise, false.</returns>
    public static bool operator !=(GimlaToDocument? left, GimlaToDocument? right)
    {
        return !(left == right);
    }

    /// <summary>
    /// Compares the current instance with another object of the same type and returns an integer that indicates whether the current instance precedes, follows, or occurs in the same position in the sort order as the other object.
    /// </summary>
    /// <param name="other">An object to compare with this instance.</param>
    /// <returns>A value that indicates the relative order of the objects being compared.</returns>
    public int CompareTo(GimlaToDocument? other)
    {
        if (other == null) return 1;

        int result = GimlaCode.CompareTo(other.GimlaCode);
        if (result == 0)
        {
            result = DocType.CompareTo(other.DocType);
        }

        return result;
    }
    #endregion
}
