using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;   // Needed for the Table class

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1, Row 1");
        builder.InsertCell();
        builder.Write("Cell 2, Row 1");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 1, Row 2");
        builder.InsertCell();
        builder.Write("Cell 2, Row 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Render the document (which contains only the table) to a PNG image.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
        // Render only the first page where the table resides.
        options.PageSet = new PageSet(0);

        string outputPath = "TableImage.png";
        doc.Save(outputPath, options);

        // Verify that the image file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the table image.");
    }
}
