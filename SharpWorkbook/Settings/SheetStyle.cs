using SharpWorkbook.Abstractions;
using SharpWorkbook.Enumerations;
using SharpWorkbook.Fonts;

namespace SharpWorkbook.Settings;

public record class SheetStyle
{
    public Alignment? Align { get; set; } = null;
    public IFontFamily? FontFamily { get; set; } = TimesNewRomanFont.Instance;
    public float FontSize { get; set; } = 12;
    public bool FontBold { get; set; } = false;
    public static SheetStyle DefaultHeader
    {
        get
        {
            return new SheetStyle()
            {
                FontBold = true,
            };
        }
    }
    public static SheetStyle DefaultBody
    {
        get
        {
            return new SheetStyle();
        }
    }
    public static SheetStyle DefaultOrdinal
    {
        get
        {
            return new SheetStyle();
        }
    }
    public static SheetStyle Default
    {
        get
        {
            return new SheetStyle();
        }
    }
}