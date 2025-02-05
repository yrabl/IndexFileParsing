namespace XmlToExcel.Objects;

public class GimlaType : IEquatable<GimlaType?>, IComparable<GimlaType>
{
    #region Members
    public int Code { get; set; }
    public string Description { get; set; } = null!;
    #endregion

    #region Methods
    public override bool Equals(object? obj)
    {
        return obj is GimlaType other && Equals(other);
    }

    public bool Equals(GimlaType? other)
    {
        return other is not null &&
               Code == other.Code;
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Code);
    }

    public static bool operator ==(GimlaType? left, GimlaType? right)
    {
        return EqualityComparer<GimlaType>.Default.Equals(left, right);
    }

    public static bool operator !=(GimlaType? left, GimlaType? right)
    {
        return !(left == right);
    }

    public int CompareTo(GimlaType? other)
    {
        if (other is null) return 1;
        return Code.CompareTo(other.Code);
    }
    #endregion
}
