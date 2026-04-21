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

        // Start a table and add a single row with three cells.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.EndRow();

        // Set the table's preferred width to 100% of the page width.
        table.PreferredWidth = PreferredWidth.FromPercent(100);

        // Finish the table.
        builder.EndTable();

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TablePreferredWidth.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output document was not created.");
        }
    }
}
