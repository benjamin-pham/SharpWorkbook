namespace SharpWorkbook.Settings;

public record WorksheetSettings
{
    public bool HasHeader { get; set; } = true;
    public bool HasOrdinal { get; set; } = true;
    public string OrdinalTitle { get; set; } = "Stt";
    public string TrueValue { get; set; } = "X";
    public string FalseValue { get; set; } = "";
    public SheetStyle SheetStyleDefault { get; set; } = SheetStyle.Default;
    public SheetStyle HeaderStyle { get; set; } = SheetStyle.DefaultHeader;
    public SheetStyle OrdinalStyle { get; set; } = SheetStyle.DefaultOrdinal;
    public SheetStyle BodyStyle { get; set; } = SheetStyle.DefaultBody;
    public static WorksheetSettings Default
    {
        get
        {
            return new WorksheetSettings();
        }
    }
}

