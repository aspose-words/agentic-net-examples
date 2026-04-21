using System;
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

        // Start a table and add a few rows/cells so the table is valid.
        Table table = builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Cell A1");
        builder.InsertCell();
        builder.Write("Cell A2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Set the table preferred width to 50 % of the page width.
        table.PreferredWidth = PreferredWidth.FromPercent(50);

        // Ensure auto‑fit is enabled (default is true). This allows columns to adjust dynamically.
        table.AllowAutoFit = true;

        // Save the document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TablePreferredWidth.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
