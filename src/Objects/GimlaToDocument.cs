
namespace XmlToExcel.Objects;

public class GimlaToDocument : IEquatable<GimlaToDocument?>, IComparable<GimlaToDocument>
{
    #region Members
    public int GimlaCode { get; set; }
    public int DocType { get; set; }
    #endregion

    #region Methods
    public override bool Equals(object? obj)
    {
        return Equals(obj as GimlaToDocument);
    }

    public bool Equals(GimlaToDocument? other)
    {
        return other is not null &&
               GimlaCode == other.GimlaCode &&
               DocType == other.DocType;
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(GimlaCode, DocType);
    }

    public static bool operator ==(GimlaToDocument? left, GimlaToDocument? right)
    {
        return EqualityComparer<GimlaToDocument>.Default.Equals(left, right);
    }

    public static bool operator !=(GimlaToDocument? left, GimlaToDocument? right)
    {
        return !(left == right);
    }

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
