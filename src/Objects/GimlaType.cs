namespace XmlToExcel.Objects;

/// <summary>
/// Represents a Gimla type with a code and description, and associated document types.
/// </summary>
public class GimlaType : IEquatable<GimlaType?>, IComparable<GimlaType>
{
    #region Members
    /// <summary>
    /// Gets or sets the code of the Gimla type.
    /// </summary>
    public int Code { get; set; }

    /// <summary>
    /// Gets or sets the description of the Gimla type.
    /// </summary>
    public string Description { get; set; } = null!;

    /// <summary>
    /// Gets the set of associated document types.
    /// </summary>
    [Browsable(false)]
    public SortedSet<DocumentType> DocumentTypes { get; } = new();
    #endregion

    #region Methods
    /// <summary>
    /// Determines whether the specified object is equal to the current object.
    /// </summary>
    /// <param name="obj">The object to compare with the current object.</param>
    /// <returns>true if the specified object is equal to the current object; otherwise, false.</returns>
    public override bool Equals(object? obj)
    {
        return obj is GimlaType other && Equals(other);
    }

    /// <summary>
    /// Determines whether the specified <see cref="GimlaType"/> is equal to the current <see cref="GimlaType"/>.
    /// </summary>
    /// <param name="other">The <see cref="GimlaType"/> to compare with the current <see cref="GimlaType"/>.</param>
    /// <returns>true if the specified <see cref="GimlaType"/> is equal to the current <see cref="GimlaType"/>; otherwise, false.</returns>
    public bool Equals(GimlaType? other)
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
    /// Determines whether two specified instances of <see cref="GimlaType"/> are equal.
    /// </summary>
    /// <param name="left">The first <see cref="GimlaType"/> to compare.</param>
    /// <param name="right">The second <see cref="GimlaType"/> to compare.</param>
    /// <returns>true if the two <see cref="GimlaType"/> instances are equal; otherwise, false.</returns>
    public static bool operator ==(GimlaType? left, GimlaType? right)
    {
        return EqualityComparer<GimlaType>.Default.Equals(left, right);
    }

    /// <summary>
    /// Determines whether two specified instances of <see cref="GimlaType"/> are not equal.
    /// </summary>
    /// <param name="left">The first <see cref="GimlaType"/> to compare.</param>
    /// <param name="right">The second <see cref="GimlaType"/> to compare.</param>
    /// <returns>true if the two <see cref="GimlaType"/> instances are not equal; otherwise, false.</returns>
    public static bool operator !=(GimlaType? left, GimlaType? right)
    {
        return !(left == right);
    }

    /// <summary>
    /// Compares the current instance with another object of the same type and returns an integer that indicates whether the current instance precedes, follows, or occurs in the same position in the sort order as the other object.
    /// </summary>
    /// <param name="other">An object to compare with this instance.</param>
    /// <returns>A value that indicates the relative order of the objects being compared.</returns>
    public int CompareTo(GimlaType? other)
    {
        if (other is null) return 1;
        return Code.CompareTo(other.Code);
    }

    /// <summary>
    /// Gets a sorted set of <see cref="GimlaToDocument"/> objects that map the Gimla type to its associated document types.
    /// </summary>
    /// <returns>A sorted set of <see cref="GimlaToDocument"/> objects.</returns>
    public SortedSet<GimlaToDocument> GetGimlaToDocuments()
    {
        var gimlaToDocuments = new SortedSet<GimlaToDocument>();
        foreach (var docType in DocumentTypes)
        {
            gimlaToDocuments.Add(new GimlaToDocument { GimlaCode = Code, GimlaDescription = Description, DocType = docType.Code, DocDescription = docType.Description });
        }
        return gimlaToDocuments;
    }
    #endregion
}
