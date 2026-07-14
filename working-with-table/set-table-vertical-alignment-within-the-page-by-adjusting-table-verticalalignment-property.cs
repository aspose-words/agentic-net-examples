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

        // Build a simple 2x2 table.
        DocumentBuilder builder = new DocumentBuilder(doc);
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

        // Finish the table and obtain the Table object.
        table = builder.EndTable();

        // Set the vertical alignment of the table on the page to Center.
        // For floating tables use RelativeVerticalAlignment; Inline tables do not support vertical alignment.
        table.RelativeVerticalAlignment = VerticalAlignment.Center;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableVerticalAlignment.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
