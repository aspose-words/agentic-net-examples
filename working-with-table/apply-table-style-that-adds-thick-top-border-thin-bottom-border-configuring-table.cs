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

        // Build a simple 2‑column table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");

        // Configure the style's borders: thick top, thin bottom.
        customStyle.Borders.Top.LineStyle = LineStyle.Single;
        customStyle.Borders.Top.LineWidth = 4.0;          // Thick top border
        customStyle.Borders.Top.Color = Color.Black;

        customStyle.Borders.Bottom.LineStyle = LineStyle.Single;
        customStyle.Borders.Bottom.LineWidth = 1.0;       // Thin bottom border
        customStyle.Borders.Bottom.Color = Color.Black;

        // Apply the style to the table.
        table.Style = customStyle;

        // Save the document.
        doc.Save("TableStyleBorders.docx");
    }
}
