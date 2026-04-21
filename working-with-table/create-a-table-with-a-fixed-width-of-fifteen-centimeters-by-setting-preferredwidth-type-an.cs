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

        // Start a table and insert the first cell (a table must have at least one cell before formatting).
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Sample cell");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Set the table's preferred width to 15 centimeters.
        // 1 inch = 2.54 cm, 1 point = 1/72 inch.
        // points = (centimeters / 2.54) * 72
        double centimeters = 15.0;
        double points = (centimeters / 2.54) * 72.0;
        table.PreferredWidth = PreferredWidth.FromPoints(points);

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableFixedWidth.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Optionally, inform that the process completed (no interactive prompts required).
        Console.WriteLine("Document created successfully: " + outputPath);
    }
}
