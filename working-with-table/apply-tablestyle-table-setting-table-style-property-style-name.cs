using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document and a DocumentBuilder to construct its contents.
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

        // Add a custom table style to the document's style collection.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
        customStyle.Borders.Color = Color.Blue;
        customStyle.Borders.LineStyle = LineStyle.Single;
        customStyle.Shading.BackgroundPatternColor = Color.LightYellow;
        customStyle.CellSpacing = 5;
        customStyle.BottomPadding = 10;
        customStyle.TopPadding = 10;
        customStyle.LeftPadding = 5;
        customStyle.RightPadding = 5;

        // Apply the style to the table via the Style property.
        table.Style = customStyle; // Directly assign the TableStyle object.
        // Alternatively, you could set the style by name:
        // table.StyleName = "MyCustomTableStyle";

        // Save the document to disk.
        doc.Save("TableWithStyle.docx");
    }
}
