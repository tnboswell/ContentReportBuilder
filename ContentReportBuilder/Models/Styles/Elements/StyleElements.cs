
using System.Data;
using System.Drawing;
/// <summary>
/// Base style elements used to build style interfaces
/// </summary>
namespace ContentReportBuilder.Models.Styles
{
    public enum HorizontalAlignment
    {
        Centre,
        Left,
        Right,
        Justified
    }

    public enum VeritcalAlignment
    {
        Top,
        Middle,
        Bottom
    }

    public class IndentationStyle
    {
        public IndentationStyle(double firstLine = 0.0, double hanging = 0.0)
        {
            FirstLine = firstLine;
            Hanging = hanging;
        }
        public double FirstLine { get; set; } // mm
        public double Hanging { get; set; } // mm

    }

    public class SpacingStyle
    {
        public SpacingStyle(double top = 0.0, double bottom = 0.0, double lineSpacing = 0.0)
        {
            Top = top;
            Bottom = bottom;
            LineSpacing = lineSpacing;
        }

        public double Top { get; set; } // pt
        public double Bottom { get; set; } // pt
        public double LineSpacing { get; set; }
    }

    public abstract class ColourConvert
    {
        public Color Colour { get; set; }
        public string HexConverter()
        {
            return "#" + Colour.R.ToString("X2") + Colour.G.ToString("X2") + Colour.B.ToString("X2");
        }
    }

    public class FontStyle: ColourConvert
    {
        public FontStyle(string font = "Arial", int size = 11)
        {
            Font = font;
            Size = size;
            Colour = Color.Black;
            IsBold = false;
            IsItalics = false;
            IsUnderlined = false;
            IsStrikeThrough = false;
        }
        public string Font { get; set; }
        public int Size { get; set; }
        
        public bool IsBold { get; set; }
        public bool IsItalics { get; set; }
        public bool IsUnderlined { get; set; }
        public bool IsStrikeThrough { get; set; }

        
    }

   

    public class CellMergingStyle
    {
        public CellMergingStyle(int columnsOffset = 0, int rowsOffset = 0)
        {
            Rows = rowsOffset;
            Columns = columnsOffset;
        }
        public int Rows { get; set; }
        public int Columns { get; set; }

        public int ColumnsMerge(int currentColumn)
        {
            return currentColumn + Columns;
        }
        public int RowsMerge(int currentRow)
        {
            return currentRow + Rows;
        }
    }

    public enum CellBorder
    {
        None,
        Thin,
        Dashed,
        Dotted,
        Medium,
        Thick,
    }
    public class CellBorderStyle : ColourConvert
    {
        public CellBorderStyle()
        {
            Top = CellBorder.None;
            Bottom = CellBorder.None;
            Left = CellBorder.None;
            Right = CellBorder.None;
            Colour = Color.Black;
        }
        public void SetAll(CellBorder type)
        {
            Top = type;
            Bottom = type;
            Left = type;
            Right = type;
        }

        public void SetAll(CellBorder type, Color colour)
        {
            Top = type;
            Bottom = type;
            Left = type;
            Right = type;
            Colour = colour;
        }

        public CellBorder Top { get; set; }
        public CellBorder Bottom { get; set; }
        public CellBorder Left { get; set; }
        public CellBorder Right { get; set; }
    }

    public class CellStyle: ColourConvert
    {
        public CellStyle(CellMergingStyle cellMerging = null, int rowOffset = 1, int columnStart = 0, CellBorderStyle border = null)
        {
            CellMerging = cellMerging;
            RowOffset = rowOffset;
            ColumnStart = columnStart;
            Colour = Color.White;
            IsWrapped = true;
            DataType = CellDataType.String;
            FillSolid = false;
            if(border == null)
                Border = new CellBorderStyle();
        }

        public CellDataType DataType { get; set; }

        public bool IsWrapped { get; set; }

        public CellBorderStyle Border { get; set; }

        public bool FillSolid { get; set; }

        /// <value>
        /// This is the row offset from the last element
        /// Default is the advance one row below the previous element
        /// </value>
        protected int RowOffset { get; set; }

        /// <value>
        /// This is the column starting location for the element. 
        /// 0 = reset to column 1
        /// any other number will offset from the end of the used cells in the worksheet.
        /// </value>
        protected int ColumnStart { get; set; }

        public int NextRow(int endRow)
        {
            if((endRow + RowOffset) < 1)
                return 1;

            return endRow + RowOffset;
        }

        public int NextColumn(int endColumn = 1)
        {
            if (ColumnStart == 0)
                return 1;

            return endColumn + ColumnStart;
        }

        public double Height = 15;
        public double Width = 8.43;
        // not null if cell merging is to be applied
        public CellMergingStyle CellMerging { get; set; }
    }

    public enum CellDataType
    {
        String,
        Integer,
        Double,
        Date
    }

    public class MarginStyle
    {
        public MarginStyle(double top = 30, double bottom = 30, double right = 25, double left = 25)
        {
            Top = top;
            Bottom = bottom;
            Left = left;
            Right = right;
        }

        public MarginStyle(double top, double bottom)
        {
            Top = top;
            Bottom = bottom;
            Left = 0;
            Right = 0;
        }

        public double Top { get; set; }//mm
        public double Bottom { get; set; }//mm
        public double Right { get; set; }//mm
        public double Left { get; set; }//mm
    }
}
