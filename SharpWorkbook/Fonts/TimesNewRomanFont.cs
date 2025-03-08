using SharpWorkbook.Abstractions;

namespace SharpWorkbook.Fonts;

internal class TimesNewRomanFont : IFontFamily
{
    public static IFontFamily Instance => new TimesNewRomanFont();

    public string Name { get { return "Times New Roman"; } }
}
