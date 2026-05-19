using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
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
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "LightGrayStyle");
        // Apply a light gray background to all cells via the style's shading.
        tableStyle.Shading.BackgroundPatternColor = Color.LightGray;

        // Assign the style to the table.
        table.Style = tableStyle;
        // Ensure the style is applied to the whole table (no conditional formatting).
        table.StyleOptions = TableStyleOptions.None;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableStyleShading.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
