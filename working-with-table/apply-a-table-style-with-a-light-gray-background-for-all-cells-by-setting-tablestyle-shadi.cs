using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableStyleShading.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Create a custom table style and set its shading to light gray.
        TableStyle grayStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "GrayTableStyle");
        grayStyle.Shading.BackgroundPatternColor = Color.LightGray;

        // Apply the style to the table.
        table.Style = grayStyle;

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
