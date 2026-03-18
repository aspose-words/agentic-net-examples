using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "ThickOuterThinInner");
        // Configure the style's default cell borders (inner borders) to be thin.
        tableStyle.Borders.LineStyle = LineStyle.Single;
        tableStyle.Borders.LineWidth = 1.0; // thin inner border
        tableStyle.Borders.Color = Color.Black;

        // Apply the style to the table.
        table.Style = tableStyle;

        // Set the outer borders of the table to be thick.
        // The last parameter (isOverrideCellBorders) is false so inner cell borders remain unchanged.
        table.SetBorder(BorderType.Left,   LineStyle.Single, 4.0, Color.Black, false);
        table.SetBorder(BorderType.Right,  LineStyle.Single, 4.0, Color.Black, false);
        table.SetBorder(BorderType.Top,    LineStyle.Single, 4.0, Color.Black, false);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 4.0, Color.Black, false);

        // Save the document.
        doc.Save("ThickOuterThinInnerTable.docx");
    }
}
