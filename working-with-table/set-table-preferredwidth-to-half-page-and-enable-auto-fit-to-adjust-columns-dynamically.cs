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

        // Start building a table.
        Table table = builder.StartTable();

        // First row with three cells.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // Second row with sample data.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.InsertCell();
        builder.Write("Row 1, Cell 3");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Set the table's preferred width to 50 % of the page width.
        table.PreferredWidth = PreferredWidth.FromPercent(50);

        // Ensure auto‑fit is enabled so columns can adjust dynamically.
        table.AllowAutoFit = true;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TablePreferredWidth.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
