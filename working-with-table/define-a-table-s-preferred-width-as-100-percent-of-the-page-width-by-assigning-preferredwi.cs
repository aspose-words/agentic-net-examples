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

        // Add a single row with three cells.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Set the table's preferred width to 100 % of the page width.
        table.PreferredWidth = PreferredWidth.FromPercent(100);

        // Validate that the preferred width was applied correctly.
        if (table.PreferredWidth.Type != PreferredWidthType.Percent ||
            Math.Abs(table.PreferredWidth.Value - 100) > 0.001)
        {
            throw new InvalidOperationException("Table preferred width is not set to 100 %.");
        }

        // Prepare the output folder and file name.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TablePreferredWidth.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("Document was not saved successfully.", outputPath);
        }
    }
}
