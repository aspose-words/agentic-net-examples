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

        // Start a table.
        Table table = builder.StartTable();

        // Insert a single cell with sample text.
        builder.InsertCell();
        builder.Write("Sample cell");

        // End the table.
        builder.EndTable();

        // Set the table's preferred width to 100% of the page width.
        table.PreferredWidth = PreferredWidth.FromPercent(100);

        // Prepare the output folder.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TablePreferredWidth.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");
    }
}
