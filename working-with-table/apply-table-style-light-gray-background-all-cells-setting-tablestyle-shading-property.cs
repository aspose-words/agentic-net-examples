using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document.
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
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "LightGrayStyle");
        // Apply a light gray background to all cells via the style's shading.
        tableStyle.Shading.BackgroundPatternColor = Color.LightGray;

        // Assign the style to the table.
        table.Style = tableStyle;

        // Save the document.
        doc.Save("TableWithLightGrayStyle.docx");
    }
}
