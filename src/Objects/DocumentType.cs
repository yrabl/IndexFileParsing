namespace XmlToExcel.Objects;

public class DocumentType : IEquatable<DocumentType?>, IComparable<DocumentType>
{
    #region Members
    public int Code { get; set; }

    public string Description { get; set; } = null!;
    #endregion

    #region Methods
    public override bool Equals(object? obj)
    {
        return Equals(obj as DocumentType);
    }

    public bool Equals(DocumentType? other)
    {
        return other is not null &&
               Code == other.Code;
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Code);
    }

    public static bool operator ==(DocumentType? left, DocumentType? right)
    {
        return EqualityComparer<DocumentType>.Default.Equals(left, right);
    }

    public static bool operator !=(DocumentType? left, DocumentType? right)
    {
        return !(left == right);
    }

    public int CompareTo(DocumentType? other)
    {
        if (other == null) return 1;
        return Code.CompareTo(other.Code);
    }
    #endregion
}
