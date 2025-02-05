
using System.ComponentModel;

namespace XmlToExcel.Objects;

public class GimlaToDocument : IEquatable<GimlaToDocument?>, IComparable<GimlaToDocument>
{
    #region Members
    [DisplayName("Gimla Code")]
    public int GimlaCode { get; set; }
    [DisplayName("Gimla Description")]
    public string GimlaDescription { get; set; } = null!;
    [DisplayName("Doc Type")]
    public int DocType { get; set; }
    [DisplayName("Doc Description")]
    public string DocDescription { get; set; } = null!;
    #endregion

    #region Methods
    public override bool Equals(object? obj)
    {
        return obj is GimlaToDocument other && Equals(other);
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
