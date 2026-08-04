using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

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
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Align the table vertically to the center of the page.
        // The RelativeVerticalAlignment property positions a floating table relative to the page.
        table.RelativeVerticalAlignment = VerticalAlignment.Center;

        // Optionally center the table horizontally as well.
        table.RelativeHorizontalAlignment = HorizontalAlignment.Center;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableVerticalAlignment.docx");
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");
    }
}
