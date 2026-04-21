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

        // Build a simple 1x1 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndTable();

        // Align the table vertically to the center of the page.
        // For floating tables this is controlled by RelativeVerticalAlignment.
        table.RelativeVerticalAlignment = VerticalAlignment.Center;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableVerticalAlignment.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
