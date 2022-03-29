using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolReport
{
    public class ExcelTools
    {
        public static SLStyle CreateAlignCenterStyle(SLDocument sl)
        {
            SLStyle textCenterStyle = sl.CreateStyle();
            textCenterStyle.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center);
            textCenterStyle.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
            return textCenterStyle;
        }
        public static SLStyle CreateTableHeaderStyle(SLDocument sl)
        {
            SLStyle tableHeaderStyle = sl.CreateStyle();
            tableHeaderStyle.Font.Bold = true;
            tableHeaderStyle.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
            tableHeaderStyle.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center);
            return tableHeaderStyle;
        }
        public static SLStyle CreateAllBorderStyle(SLDocument sl)
        {
            SLStyle allBorderStyle = sl.CreateStyle();
            allBorderStyle.Border.BottomBorder.BorderStyle = BorderStyleValues.Thin;
            allBorderStyle.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
            allBorderStyle.Border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
            allBorderStyle.Border.TopBorder.BorderStyle = BorderStyleValues.Thin;
            return allBorderStyle;
        }
        public static SLStyle CreateHeader2Style(SLDocument sl)
        {
            SLStyle header2Type = sl.CreateStyle();
            header2Type.SetFontBold(true);
            header2Type.Font.FontSize = 16;
            header2Type.Font.FontName = "Calibri";
            return header2Type;
        }

        public static SLStyle CreateHeader1Style(SLDocument sl)
        {
            SLStyle header1Type = sl.CreateStyle();


            header1Type.SetFontBold(true);
            header1Type.Font.FontSize = 24;
            header1Type.Font.FontName = "Calibri";
            header1Type.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
            header1Type.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center);

            return header1Type;
        }
        public static SLStyle SetColorStyle(SLDocument sl, System.Drawing.Color color)
        {
            SLStyle style = sl.CreateStyle();
            style.Fill.SetPattern(PatternValues.Solid, color, color);
            return style;
        }

    }
}
