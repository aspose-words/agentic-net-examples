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

        // Start a new table.
        Table table = builder.StartTable();

        // Insert first row with two cells.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Insert second row with two cells.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Set a fixed width for the table (e.g., 300 points) and center it on the page.
        table.PreferredWidth = PreferredWidth.FromPoints(300);
        table.Alignment = TableAlignment.Center;

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableFixedWidthCentered.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Optionally, inform that the process completed (no interactive input required).
        Console.WriteLine("Document created successfully at: " + outputPath);
    }
}
