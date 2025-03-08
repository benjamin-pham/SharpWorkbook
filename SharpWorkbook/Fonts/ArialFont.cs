using SharpWorkbook.Abstractions;

namespace SharpWorkbook.Fonts;

internal class ArialFont : IFontFamily
{
    public static IFontFamily Instance => new ArialFont();

    public string Name { get { return "Arial"; } }
}