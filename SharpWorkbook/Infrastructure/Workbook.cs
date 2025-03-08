using ClosedXML.Excel;
using SharpWorkbook.Abstractions;
using SharpWorkbook.ColumnCustoms;
using SharpWorkbook.Enumerations;
using SharpWorkbook.Settings;

namespace SharpWorkbook.Infrastructure;

internal class Workbook : IWorkbook
{
    private readonly IXLWorkbook _workbook;
    private MemoryStream _stream = new MemoryStream();

    public Workbook()
    {
        _workbook = new XLWorkbook();
    }

    public void AddSheet<T>(
        string sheetName,
        IList<T> data,
        IList<IWorksheetColumn>? columns = null,
        WorksheetSettings? settings = null)
        where T : class
    {
        new Worksheet<T>(WorksheetDataset<T>.Create(sheetName, data, settings, columns), _workbook.AddWorksheet(sheetName)).Handle();
    }
    public void AddSheet<T>(WorksheetDataset<T> dataset) where T : class
    {
        new Worksheet<T>(dataset, _workbook.AddWorksheet(dataset.SheetName)).Handle();
    }
    public byte[] GetBytes()
    {
        _workbook.SaveAs(_stream);
        _stream.Position = 0;
        return _stream.ToArray();
    }

    public void RemoveSheet(string sheetName)
    {
        _workbook.Worksheets.Delete(sheetName);
    }

    public void SaveAs(string file)
    {
        _workbook.SaveAs(file);
    }

    internal class Worksheet<T>
    {
        private readonly IXLWorksheet _worksheet;

        private object[,] _data2d;

        private SheetStyle _headerStyle;

        private SheetStyle _ordinalStyle;

        private SheetStyle _bodyStyle;

        private int _rowsCount;

        private int _columnsCount;

        private bool _hasHeader;

        private IList<IWorksheetColumn>? _columns;

        private int _addOrdinal;
        private int _addHeader;
        public Worksheet(WorksheetDataset<T> dataset, IXLWorksheet worksheet)
        {
            _worksheet = worksheet;
            _data2d = dataset.Data2d;
            _headerStyle = dataset.Settings!.HeaderStyle;
            _ordinalStyle = dataset.Settings!.OrdinalStyle;
            _bodyStyle = dataset.Settings!.BodyStyle;
            _rowsCount = dataset.RowsCount;
            _columnsCount = dataset.ColumnsCount;
            _hasHeader = dataset.Settings.HasHeader;
            _columns = dataset.Columns;
            _addOrdinal = dataset.Settings.HasOrdinal ? 1 : 0;
            _addHeader = dataset.Settings.HasHeader ? 1 : 0;            
        }

#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
        public Worksheet() { }
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.

        public void Handle()
        {
            LoadSheet();

            SetStyle();
        }

        private void LoadSheet()
        {
            for (int row = 0; row < _rowsCount; row++)
            {
                for (int col = 0; col < _columnsCount; col++)
                {
                    _worksheet.Cell(row + 1, col + 1).Value = _data2d[row, col].ToString();
                }
            }
            if (_hasHeader)
                _worksheet.Cell(1, 1).Value = "STT";
        }

        public void SetStyle()
        {
            SetHeaderStyle();
            SetOrdinalStyle();
            SetBodyStyle();
            AdjustColumn();
        }

        public void SetHeaderStyle()
        {
            if (!_hasHeader)
                return;

            IXLRange headerRange = _worksheet.Range(1, 1, 1, _columnsCount);

            SetFont(headerRange, _headerStyle);
        }

        public void AdjustColumn()
        {
            for (int i = 0; i < _columnsCount; i++)
            {
                IWorksheetColumn header = _columns![i];

                IXLColumn col = _worksheet.Column(i + 1);

                if (header.Width.HasValue)
                {
                    col.Width = header.Width.Value;
                }
                else
                {
                    col.AdjustToContents();
                }
            }
        }

        public void SetOrdinalStyle()
        {

        }

        public void SetBodyStyle()
        {
            IXLCell beginBody = _worksheet.Cell(1 + _addHeader, 1 + _addOrdinal);
            IXLCell endBody = _worksheet.Cell(_data2d.GetLength(0), _data2d.GetLength(1));

            IXLRange bodyRange = _worksheet.Range(beginBody, endBody);

            SetFont(bodyRange, _bodyStyle);
        }

        private void SetFont(IXLRange range, SheetStyle stylesheet)
        {
            IXLFont font = range.Style.Font;
            font.Bold = stylesheet.FontBold;
            font.FontName = stylesheet.FontFamily?.Name;
            font.FontSize = stylesheet.FontSize;
        }

        private void SetDataFormat(IFormatCell formatCell, IXLColumn xLColumn)
        {
            switch (formatCell.DataType)
            {
                case CellDataType.Text:
                    xLColumn.AsRange().CellsUsed().setd
                    break;
                case CellDataType.Number:
                    break;
                case CellDataType.Boolean:
                    break;
                case CellDataType.DateTime:
                    break;
                case CellDataType.TimeSpan:
                    break;
                default: break;
            }
        }
    }
}
