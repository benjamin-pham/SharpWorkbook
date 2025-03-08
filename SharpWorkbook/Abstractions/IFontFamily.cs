namespace SharpWorkbook.Abstractions;

public interface IFontFamily
{
    static abstract IFontFamily Instance { get; }
    string Name { get; }
}